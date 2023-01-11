# Publish PowerShell Module

## Publish to GitHub Packages as a NuGet Package

### Pre-Requisites:

1. Install the latest `BETA` version of `PowerShellGet` from the PowerShell Gallery:

```powershell
# Ensuring PowerShellGet stable is latest version
Install-Module -Name PowerShellGet -Force -AllowClobber
# Installing PowerShellGet 3 Prerelease
Install-Module -Name PowerShellGet -RequiredVersion 3.0.16-beta16 -AllowPrerelease -Force -Repository PSGallery -SkipPublisherCheck
```

2. Install the [NuGet CLI](https://docs.microsoft.com/en-us/nuget/install-nuget-client-tools)

```powershell
winget install Microsoft.NuGet
```

3. Install the [gpr](https://github.com/jcansdale/gpr) tool:

```powershell
dotnet tool install --global gpr
```

4. Create a [Personal Access Token](https://docs.github.com/en/github/authenticating-to-github/creating-a-personal-access-token)
    with the `read:packages` and `write:packages` scopes.

```powershell
gh secret set --user GHPKG_TOKEN -b $ENV:GHPKG_TOKEN
```

### Workflow

1. Create a "local" `Nuget` Repository:

```powershell
New-Item -Path $LocalRepoPath -Type Directory | Out-Null
Register-PSResourceRepository -Name "LocalRepo" -Uri $LocalRepoPath
```

2. Publish the module to the local repository as a *PSResource*:

```powershell
Publish-PSResource -Path "./*.psd1" -Repository "LocalRepo"
```

3. Using `gpr`, publish to GitHub Packages for your repository:

```powershell
$NupkgPath = Get-ChildItem -Path $LocalRepoPath -Include "*.nupkg" -Recurse
gpr push -k $Env:GHPKG_TOKEN $NupkgPath -r "https://github.com/<user>/<repo>"
```

## Publish to PowerShell Gallery


```powershell
Publish-PSResource -Path <path/to/module> -Repository "PSGallery" -ApiKey $Env:NUGET_API_TOKEN
```

or using the `build.ps1` script:

```powershell
# setup API key
$apiKey = $Env:NUGET_API_TOKEN | ConvertTo-SecureString -AsPlainText -Force
$cred = [pscredential]::new('apikey', $apiKey)

# run build "Publish" task
./build.ps1 -Publish -PSGalleryApiKey $cred -Bootstrap 
```
