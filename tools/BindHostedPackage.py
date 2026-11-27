"""Bind a successful hosted release build to its package and consumer receipts."""

import argparse
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
import subprocess
from xml.etree import ElementTree
from zipfile import ZipFile


REPOSITORY = Path(__file__).resolve().parent.parent
API = "repos/Lokad/OoxPDF/actions"
FRAMEWORKS = {"net11.0": "ci-candidate", "net10.0": "ci-candidate-net10"}
JOBS = {"verify (windows-latest)", "verify (ubuntu-latest)", "Package consumer (ubuntu-latest)"}
SUMMARY = re.compile(r"package-smoke/[0-9a-f]{32}/summary\.json")
RUNTIME = re.compile(r"ci-candidate(?:-net10)?/(?:identity\.json|src/Lokad\.OoxPdf\.Cli/bin/Release/"
                     r"net1[01]\.0/Lokad\.OoxPdf(?:\.Cli)?\.(?:dll|exe|pdb|deps\.json|runtimeconfig\.json))")
CONSUMER_CHECKS = (
    "MemoryAdmissionChecked", "PptxConversionChecked", "CancellationChecked", "MalformedInputChecked",
    "AdmissionRefusalChecked", "FailurePreservesDestinationChecked", "VersionMetadataChecked",
    "FrameworkBuildIdentitiesChecked", "CallerDocumentationChecked", "AsyncOnlyOutputChecked",
    "ExplicitTemporaryStorageChecked", "TemporaryStorageCleanupChecked", "FontLeaseContractChecked",
)


def require(condition, detail):
    if not condition:
        raise ValueError(detail)


def read(path):
    return json.loads(path.read_bytes())


def digest(data):
    return sha256(data).hexdigest().upper()


def pin(path, data=None):
    return {"Path": path.as_posix(), "Sha256": digest(path.read_bytes() if data is None else data)}


def publish(path, data):
    require((REPOSITORY / path).resolve() == REPOSITORY / path, f"Aliased output {path}")
    if path.exists():
        require(path.is_file() and path.read_bytes() == data, f"Conflicting output {path}")
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as output:
        output.write(data)


def publish_json(path, value):
    publish(path, (json.dumps(value, indent=2) + "\n").encode("utf-8"))


def git(*args):
    return subprocess.check_output(["git", *args], cwd=REPOSITORY)


def artifact(stage, metadata, name, run, revision, offline):
    entry, = [row for row in metadata["artifacts"] if row["name"] == name]
    workflow = entry["workflow_run"]
    require(workflow["id"] == run and workflow["head_sha"] == revision, f"Wrong run for {name}")
    path = stage / (name + ".zip")
    if path.exists():
        payload = path.read_bytes()
    else:
        require(not offline and not entry["expired"], f"Verified archive unavailable for {name}")
        payload = subprocess.check_output(["gh", "api", f"{API}/artifacts/{int(entry['id'])}/zip"])
    require(len(payload) == entry["size_in_bytes"] and
            digest(payload) == entry["digest"].removeprefix("sha256:").upper(),
            f"Archive digest or length differs for {name}")
    publish(path, payload)
    return path


def zip_files(source, select):
    with ZipFile(source) as archive:
        names = [entry.filename for entry in archive.infolist() if not entry.is_dir()]
        require(len(names) == len({name.casefold() for name in names}), "Duplicate archive entries")
        return {name: archive.read(name) for name in names if select(name)}


def bind(stage, offline):
    inputs = read(stage / "bound-inputs.json")
    revision = inputs["Revision"]
    require(re.fullmatch(r"[0-9a-f]{40}", revision), "Expected a full release revision")
    trigger, = read(stage / "hosted-trigger.json")
    status_path = stage / "hosted-ci.json"
    status = read(status_path)
    run = int(trigger["databaseId"])
    require(trigger["headSha"] == status["headSha"] == revision and
            status["url"].endswith(f"/actions/runs/{run}") and
            status["status"] == "completed" and status["conclusion"] == "success",
            "Hosted run does not qualify the requested revision")
    jobs = status["jobs"]
    require(len(jobs) == len(JOBS) and {row["name"] for row in jobs} == JOBS and
            all(row["status"] == "completed" and row["conclusion"] == "success" for row in jobs),
            "Expected successful Windows, Ubuntu and package consumer jobs")
    remote = "refs/remotes/origin/master"
    if offline:
        git("merge-base", "--is-ancestor", revision, remote)
    else:
        require(git("rev-parse", remote).decode().strip() == revision, "Remote master revision differs")
    project = ElementTree.fromstring(git("show", f"{revision}:src/Lokad.OoxPdf/Lokad.OoxPdf.csproj"))
    version = project.findtext("./PropertyGroup/Version")
    require(version and re.fullmatch(r"\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?", version), "Package version differs")
    assembly_version = version.split("-", 1)[0] + ".0"

    metadata_path = stage / "hosted-artifacts.json"
    if not metadata_path.exists():
        require(not offline, "Offline binding needs cached artifact metadata")
        publish_json(metadata_path, json.loads(subprocess.check_output(["gh", "api", f"{API}/runs/{run}/artifacts"])))
    metadata = read(metadata_path)
    windows = artifact(stage, metadata, "candidate-windows-latest", run, revision, offline)
    ubuntu = artifact(stage, metadata, "package-smoke-ubuntu-latest", run, revision, offline)
    prefix = f"Lokad.OoxPdf.{version}"
    names = {f"nuget/{prefix}.nupkg", f"nuget/{prefix}.snupkg"}
    win_files = zip_files(windows, lambda name: name.endswith((".nupkg", ".snupkg")) or
                         RUNTIME.fullmatch(name) or SUMMARY.fullmatch(name))
    require({name for name in win_files if name.endswith((".nupkg", ".snupkg"))} == names,
            "Expected one package and one symbol archive")
    package_data = win_files.pop(f"nuget/{prefix}.nupkg")
    symbol_data = win_files.pop(f"nuget/{prefix}.snupkg")
    ubuntu_files = zip_files(ubuntu, SUMMARY.fullmatch)
    identities = {}
    with ZipFile(BytesIO(package_data)) as package, ZipFile(BytesIO(symbol_data)) as symbols:
        for framework, candidate in FRAMEWORKS.items():
            identity_name = f"{candidate}/identity.json"
            identity = json.loads(win_files[identity_name])
            require(identity["Revision"] == revision and identity["Framework"] == framework,
                    f"Frozen identity differs for {framework}")
            key = f"src/Lokad.OoxPdf/bin/Release/{framework}/Lokad.OoxPdf"
            cli = f"{candidate}/src/Lokad.OoxPdf.Cli/bin/Release/{framework}/Lokad.OoxPdf"
            library = package.read(f"lib/{framework}/Lokad.OoxPdf.dll")
            pdb = symbols.read(f"lib/{framework}/Lokad.OoxPdf.pdb")
            require(identity["AssemblyRevisions"][key + ".dll"] == revision and
                    digest(library) == identity["Assets"][key + ".dll"] == identity["LibrarySha256"] and
                    win_files[cli + ".dll"] == library, f"Packaged library differs for {framework}")
            require(digest(pdb) == identity["Assets"][key + ".pdb"], f"Packaged symbols differ for {framework}")
            win_files[cli + ".pdb"] = pdb
            directory = stage / "ci-windows"
            identities[framework] = {
                "Identity": pin(directory / identity_name, win_files[identity_name]),
                "Library": pin(directory / (cli + ".dll"), library),
                "Pdb": pin(directory / (cli + ".pdb"), pdb),
            }

    package_path = Path("artifacts/nuget") / f"{prefix}.source-{revision[:9]}.nupkg"
    symbol_path = package_path.with_suffix(".snupkg")
    package_pin = pin(package_path, package_data)
    consumers = {}
    for host, rows in (("Windows", win_files), ("Ubuntu", ubuntu_files)):
        summaries = [(name, json.loads(data)) for name, data in rows.items() if SUMMARY.fullmatch(name)]
        for framework in FRAMEWORKS:
            (name, receipt), = [row for row in summaries if row[1].get("Framework") == framework
                                and row[1].get("ReusedPackage") is True]
            library_hash = identities[framework]["Library"]["Sha256"]
            require(receipt["LibraryRebuilt"] is False and all(receipt[key] is True for key in CONSUMER_CHECKS)
                    and receipt["PackageVersion"] == version and
                    receipt["PackageSha256"] == receipt["RestoredPackageSha256"] == package_pin["Sha256"] and
                    receipt["LibrarySha256"] == receipt["ExpectedLibrarySha256"] == library_hash,
                    f"Consumer evidence differs for {host}/{framework}")
            entries = {row["Entry"]: row for row in receipt["FrameworkIdentities"]}
            require(len(entries) == len(receipt["FrameworkIdentities"]) == len(FRAMEWORKS),
                    "Expected distinct identities for both frameworks")
            for target in FRAMEWORKS:
                info = entries[f"lib/{target}/Lokad.OoxPdf.dll"]
                require(info["InformationalVersion"] == f"{version}+{revision}" and
                        info["AssemblyVersion"] == info["FileVersion"] == assembly_version and
                        info["Sha256"] == identities[target]["Library"]["Sha256"],
                        f"Consumer framework identity differs for {host}/{target}")
            consumers[f"{host}/{framework}"] = pin(stage / f"ci-{host.lower()}" / name, rows[name])

    publish(package_path, package_data)
    publish(symbol_path, symbol_data)
    for host, rows in (("windows", win_files), ("ubuntu", ubuntu_files)):
        for name, data in rows.items():
            publish(stage / f"ci-{host}" / name, data)
    publish_json(stage / "bound-hosted.json", {
        "Revision": revision, "Run": run, "CI": pin(status_path),
        "Package": package_pin, "Symbols": pin(symbol_path, symbol_data), "Identities": identities,
        "Consumers": consumers, "Archives": {"Windows": pin(windows), "Ubuntu": pin(ubuntu)},
        "Metadata": pin(metadata_path),
    })
    print("Bound the package and symbols to both framework libraries and four hosted consumers.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", type=Path, help="Existing qualification directory under artifacts/")
    parser.add_argument("--offline", action="store_true", help="Replay retained archives for a revision on origin/master")
    args = parser.parse_args()
    require(Path.cwd() == REPOSITORY, "Run from the repository root")
    directory = args.stage.absolute()
    require(directory == directory.resolve(strict=True) and directory.is_dir() and
            (REPOSITORY / "artifacts") in directory.parents, "Expected a physical qualification directory")
    bind(directory.relative_to(REPOSITORY), args.offline)


if __name__ == "__main__":
    main()
