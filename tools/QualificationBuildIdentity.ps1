function Get-QualificationProductionInputs {
    @('src', 'tools/Lokad.OoxPdf.DocxInspect',
        ':(top,glob)**/*.props', ':(top,glob)**/*.targets', ':(top,glob)**/*.csproj',
        'Directory.Build.rsp', 'tools/Directory.Build.rsp', 'global.json', 'tools/global.json',
        'NuGet.Config', 'nuget.config', 'tools/NuGet.Config', 'tools/nuget.config', '.gitattributes')
}

function Assert-QualificationProductionInputs([string] $Repository, [string] $Before, [string] $After) {
    $buildInputs = @(Get-QualificationProductionInputs)
    $changed = @(git -C $Repository diff --name-only --no-ext-diff $Before $After -- @buildInputs)
    if ($LASTEXITCODE -ne 0) { throw 'Cannot verify qualification production inputs' }
    if ($changed.Count) { throw 'Production source or build inputs changed; build a new runtime' }
}

function Read-ManagedAssemblyIdentity([IO.Stream] $Stream) {
    $pe = [System.Reflection.PortableExecutable.PEReader]::new($Stream,
        [System.Reflection.PortableExecutable.PEStreamOptions]::LeaveOpen)
    try {
        $reader = [System.Reflection.Metadata.PEReaderExtensions]::GetMetadataReader($pe)
        $assembly = $reader.GetAssemblyDefinition()
        $versions = @{}
        foreach ($handle in $assembly.GetCustomAttributes()) {
            $attribute = $reader.GetCustomAttribute($handle)
            if ($attribute.Constructor.Kind -ne [System.Reflection.Metadata.HandleKind]::MemberReference) { continue }
            $member = $reader.GetMemberReference([System.Reflection.Metadata.MemberReferenceHandle]::op_Explicit($attribute.Constructor))
            if ($member.Parent.Kind -ne [System.Reflection.Metadata.HandleKind]::TypeReference) { continue }
            $type = $reader.GetTypeReference([System.Reflection.Metadata.TypeReferenceHandle]::op_Explicit($member.Parent))
            $name = $reader.GetString($type.Name)
            if ($reader.GetString($type.Namespace) -cne 'System.Reflection' -or
                $name -cnotin @('AssemblyFileVersionAttribute', 'AssemblyInformationalVersionAttribute')) { continue }
            $value = $reader.GetBlobReader($attribute.Value)
            if ($value.ReadUInt16() -ne 1) { throw 'Invalid managed assembly version metadata' }
            $versions[$name] = $value.ReadSerializedString()
        }
        $Stream.Position = 0
        return @{ AssemblyVersion = $assembly.Version.ToString()
            FileVersion = $versions.AssemblyFileVersionAttribute
            InformationalVersion = $versions.AssemblyInformationalVersionAttribute
            Sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Stream)) }
    }
    finally { $pe.Dispose() }
}

function Read-QualificationAssemblyIdentity([string] $Path) {
    $stream = [IO.File]::OpenRead($Path)
    try {
        $identity = Read-ManagedAssemblyIdentity $stream
        if ($identity.InformationalVersion -cnotmatch '[+.](?<revision>[0-9a-f]{40}|[0-9a-f]{64})\z') {
            throw ('Release assembly does not carry an SDK source revision: ' + $Path)
        }
        return @{ Revision = $Matches.revision; Sha256 = $identity.Sha256 }
    }
    finally { $stream.Dispose() }
}
