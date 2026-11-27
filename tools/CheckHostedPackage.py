"""Run hosted package transfer controls or verify a staged package."""

import argparse
import os
from pathlib import Path
import re
import subprocess
import sys
from zipfile import ZipFile

from BindHostedPackage import FRAMEWORKS, REPOSITORY, digest, git, pin, publish_json, read, require


HOSTS = ("windows", "ubuntu")
NUGET = Path("artifacts/nuget")
REMOTE = "refs/remotes/origin/master"
OUTCOME_ASSETS = ("Source", "Pdf", "Diagnostics")


def physical(path):
    absolute = REPOSITORY / path
    require(absolute.is_relative_to(REPOSITORY) and absolute.resolve(strict=True) == absolute,
            f"Expected a physical repository path: {path}")
    return absolute.relative_to(REPOSITORY)


def checked(asset):
    path = physical(Path(asset["Path"]))
    require(pin(path) == asset, f"Recorded digest differs: {path}")
    return path


def load_receipt(stage, binding, path, host, phase, runtimes):
    receipt = read(physical(path))
    runtime_revision = binding["LocalRevision"] if phase == "local" else binding["Revision"]
    require((receipt["Revision"], receipt["RuntimeRevision"], receipt["Platform"], receipt["Phase"]) ==
            (binding["Revision"], runtime_revision, host, phase) and
            checked(receipt["Binding"]) == stage / "bound-inputs.json", f"Receipt identity differs: {path}")
    require(set(receipt["Checks"]) == {f"{case}-{framework}" for case in binding["Controls"] for framework in runtimes},
            f"Incomplete transfer receipt: {path}")
    for name, row in receipt["Checks"].items():
        case, framework = name.rsplit("-", 1)
        runtime, library_hash = runtimes[framework]
        control = binding["Controls"][case]
        expected = control["ExpectedWindows" if host == "windows" else "ExpectedUbuntu"]
        require(row["Runtime"] == runtime and row["LibrarySha256"] == library_hash and
                row["Source"] == control["Source"] and row["StrictExit"] in (0, 3) and
                row["StrictExit"] == expected["StrictExit"],
                f"Transfer identity differs: {path} / {name}")
        for key in OUTCOME_ASSETS:
            checked(row[key])
            checked(expected[key])
            require(row[key]["Sha256"] == expected[key]["Sha256"], f"Transfer {key} differs: {path} / {name}")
    return receipt["Checks"]


def local_runtime(binding):
    runtime = binding["LocalIdentity"]
    identity = read(checked(runtime))
    require((identity["Revision"], identity["Framework"]) == (binding["LocalRevision"], "net11.0"),
            "Local qualification runtime differs")
    return {"net11.0": (runtime, identity["LibrarySha256"])}


def hosted_runtimes(bound, revision):
    require(bound["Revision"] == revision and set(bound["Identities"]) == set(FRAMEWORKS),
            "Hosted bindings differ from the selected revision")
    status = read(checked(bound["CI"]))
    require(status["headSha"] == revision and status["status"] == "completed" and status["conclusion"] == "success",
            "Hosted CI has not qualified the selected revision")
    runtimes = {}
    for framework, assets in bound["Identities"].items():
        identity = read(checked(assets["Identity"]))
        library = checked(assets["Library"])
        require((identity["Revision"], identity["Framework"], identity["LibrarySha256"]) ==
                (revision, framework, assets["Library"]["Sha256"]) and
                library.parent == Path(assets["Identity"]["Path"]).parent /
                f"src/Lokad.OoxPdf.Cli/bin/Release/{framework}", f"Hosted runtime differs: {framework}")
        runtimes[framework] = assets["Identity"], assets["Library"]["Sha256"]
    return runtimes


def transfer(stage, dotnet):
    host = "windows" if sys.platform == "win32" else "ubuntu"
    binding_path = physical(stage / "bound-inputs.json")
    binding = read(binding_path)
    revision = binding["Revision"]
    require(git("rev-parse", REMOTE).decode().strip() == revision, "Remote master revision differs")
    runtimes = hosted_runtimes(read(physical(stage / "bound-hosted.json")), revision)
    local = load_receipt(stage, binding, stage / f"local-{host}/evidence.json", host, "local",
                         local_runtime(binding))
    require(local, "No qualified transfer controls")
    for font in binding["Fonts"]:
        checked(font)
    for name in binding["Controls"]:
        require(re.fullmatch(r"[a-z0-9][a-z0-9-]*", name), f"Invalid control name: {name}")

    output = stage / f"hosted-{host}"
    output.mkdir()
    scratch = (output / "tmp").absolute()
    scratch.mkdir()
    environment = dict(os.environ, TMP=str(scratch), TEMP=str(scratch), TMPDIR=str(scratch))
    workdir = REPOSITORY if host == "windows" else REPOSITORY / physical(Path(binding["PortableWorkdir"]))
    checks = {}
    for framework, (runtime, library_hash) in runtimes.items():
        path = checked(runtime)
        identity = read(path)
        prefix = f"src/Lokad.OoxPdf.Cli/bin/Release/{framework}"
        binary = path.parent / prefix
        for name in ("Lokad.OoxPdf.Cli.dll", "Lokad.OoxPdf.Cli.deps.json",
                     "Lokad.OoxPdf.Cli.runtimeconfig.json", "Lokad.OoxPdf.dll"):
            checked({"Path": (binary / name).as_posix(), "Sha256": identity["Assets"][f"{prefix}/{name}"]})
        for name, control in binding["Controls"].items():
            source = checked(control["Source"])
            case = f"{name}-{framework}"
            pdf = output / (case + ".pdf")
            diagnostics = pdf.with_suffix(".json")
            result = subprocess.run([dotnet, (binary / "Lokad.OoxPdf.Cli.dll").absolute(), "convert",
                                     source.absolute(), pdf.absolute(), "--strict", "--diagnostics",
                                     diagnostics.absolute(), *control["Options"]], cwd=workdir, env=environment,
                                    capture_output=True, timeout=60)
            pdf.with_suffix(".log").write_bytes(result.stdout + result.stderr)
            require(pdf.is_file() and diagnostics.is_file(), f"Hosted conversion did not produce artifacts: {host} / {case}")
            expected = local[f"{name}-net11.0"]
            row = dict(Source=pin(source), Pdf=pin(pdf), Diagnostics=pin(diagnostics), StrictExit=result.returncode,
                       Runtime=runtime, LibrarySha256=library_hash)
            require(row["StrictExit"] == expected["StrictExit"] and
                    all(row[key]["Sha256"] == expected[key]["Sha256"] for key in OUTCOME_ASSETS),
                    f"Hosted conversion differs: {host} / {case}")
            checks[case] = row
    require(pin(binding_path) == read(physical(stage / f"local-{host}/evidence.json"))["Binding"],
            "Transfer inputs changed during conversion")
    publish_json(output / "evidence.json", dict(Revision=revision, Phase="hosted", Platform=host,
                RuntimeRevision=revision, Binding=pin(binding_path), Checks=checks))
    print(f"{host}: {len(checks)} hosted transfers match their qualified controls.")


def readback(stage):
    post = read(physical(stage / "post-promotion.json"))
    promotion = read(checked(post["Promotion"]))
    qualification_path = checked(post["Qualification"])
    qualification = read(qualification_path)
    bound = read(checked(qualification["hostedBinding"]))
    binding = read(physical(stage / "bound-inputs.json"))
    revision = post["Revision"]
    require(re.fullmatch("[a-f0-9]{40}", revision) and
            post["Stage"] == promotion["checkpoint"] == qualification["checkpoint"] and
            revision == promotion["revision"] == qualification["revision"] == binding["Revision"] and
            post["Publication"] == qualification["publication"] == promotion["publication"] == "Deferred",
            "Staging receipts disagree")
    git("merge-base", "--is-ancestor", revision, REMOTE)
    runtimes = hosted_runtimes(bound, revision)

    package, symbols = checked(post["Package"]), checked(post["Symbols"])
    version = re.fullmatch(r"Lokad\.OoxPdf\.(\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)\.nupkg", package.name)
    require(version and package.parent == symbols.parent == NUGET and symbols == package.with_suffix(".snupkg"),
            "Expected canonical package and symbol names")
    checksums = f"SHA256SUMS.{version[1]}.txt"
    package_receipt = f"{package.stem}.qualification.json"
    expected_names = {p.name for p in (package, symbols, package.with_suffix(".nupkg.sha256"),
                                      symbols.with_suffix(".snupkg.sha256"))} | {checksums, package_receipt}
    require(set(promotion["updated"]) == set(promotion["rollback"]) == expected_names and
            post["FlatFiles"] == post["PreviousStageRollbackFiles"] == len(expected_names), "Flat assets differ")
    for assets in (promotion["updated"], promotion["published"]):
        for name, sha in assets.items():
            require(Path(name).name == name, f"Expected a flat asset: {name}")
            checked({"Path": (NUGET / name).as_posix(), "Sha256": sha})
    prior = read(checked(qualification["priorStage"]))
    previous = read(checked(prior["Promotion"]))
    rollback_paths = {checked(asset) for asset in promotion["rollback"].values()}
    require(len(rollback_paths) == len(expected_names) and all(path.parent == NUGET for path in rollback_paths) and
            {name: asset["Sha256"] for name, asset in promotion["rollback"].items()} == previous["updated"] and
            promotion["published"] == previous["published"] and post["PublishedFiles"] == len(promotion["published"]) and
            post["RetainedRollbackFiles"] == prior["RetainedRollbackFiles"] + len(expected_names),
            "Published asset or rollback count differs")
    checksum_parts = []
    for path, key in ((package, "Package"), (symbols, "Symbols")):
        sha = pin(path)["Sha256"]
        checked(bound[key])
        checksum = f"{sha}  {path.name}\n".encode("ascii")
        require(qualification[key.lower()] == bound[key] and sha == bound[key]["Sha256"] and
                physical(path.with_suffix(path.suffix + ".sha256")).read_bytes() ==
                checksum, f"Archive or checksum differs: {path}")
        checksum_parts.append(checksum)
    require((NUGET / checksums).read_bytes() == b"".join(checksum_parts) and
            (NUGET / package_receipt).read_bytes() == qualification_path.read_bytes(), "Canonical evidence differs")
    with ZipFile(package) as payload, ZipFile(symbols) as debug:
        for framework, assets in bound["Identities"].items():
            require(digest(payload.read(f"lib/{framework}/Lokad.OoxPdf.dll")) == assets["Library"]["Sha256"] and
                    digest(debug.read(f"lib/{framework}/Lokad.OoxPdf.pdb")) == assets["Pdb"]["Sha256"],
                    f"Staged binaries differ: {framework}")

    expected_consumers = {f"{host}/{framework}" for host in ("Windows", "Ubuntu") for framework in FRAMEWORKS}
    require(set(bound["Consumers"]) == expected_consumers and post["SameArchiveConsumers"] == len(expected_consumers),
            "Incomplete consumer verification")
    for name, asset in bound["Consumers"].items():
        receipt = read(checked(asset))
        _, framework = name.split("/")
        require(receipt["Framework"] == framework and receipt["LibrarySha256"] == runtimes[framework][1] and
                receipt["PackageSha256"] == receipt["RestoredPackageSha256"] == post["Package"]["Sha256"],
                f"Consumer used different bytes: {name}")

    baseline = read(checked(qualification["baselineInk"]))
    require(baseline["Revision"] == revision, "Local baseline revision differs")
    for threshold in read(checked(qualification["portfolio"]))["Thresholds"].values():
        checked(threshold)
    ink = {}
    local = local_runtime(binding)
    for host in HOSTS:
        load_receipt(stage, binding, checked(baseline["LocalReports"][host]), host, "local", local)
        hosted = load_receipt(stage, binding, checked(qualification["integration"][host]), host, "hosted", runtimes)
        for name, item in hosted.items():
            ink[f"{host}/{name}"] = {key: item[key] for key in ("Pdf", "Diagnostics", "StrictExit")}
    require(ink and len(ink) == post["HostedTransferPairs"], "Hosted transfer count differs")
    result = dict(Revision=revision, Package=pin(package), Symbols=pin(symbols), TransferInk=ink,
                  RollbackNames=post["RetainedRollbackFiles"], PublishedAssetCount=post["PublishedFiles"],
                  NativeEvidence=qualification["baselineInk"])
    receipt = stage / "staging-readback.json"
    if receipt.exists():
        require(read(physical(receipt)) == result, "Retained readback differs")
    else:
        publish_json(receipt, result)
    print(f"Verified the staged archive, {len(expected_consumers)} consumers and {len(ink)} transfers.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=("transfer", "readback"))
    parser.add_argument("stage", type=Path, help="Existing qualification directory under artifacts/")
    parser.add_argument("--dotnet", default=str(Path.home() / ".dotnet" / ("dotnet.exe" if os.name == "nt" else "dotnet")),
                        help="Path to the dotnet host for transfer")
    args = parser.parse_args()
    require(Path.cwd() == REPOSITORY, "Run from the repository root")
    stage = physical(args.stage)
    require(stage.is_dir() and Path("artifacts") in stage.parents, "Expected a qualification directory in artifacts/")
    if args.operation == "transfer":
        transfer(stage, args.dotnet)
    else:
        readback(stage)


if __name__ == "__main__":
    main()
