param(
    [string]$BatDir = $null,
    [string]$Head = $null,
    [string]$Base = $null,
    [string]$Title = $null,
    [string]$NewBranch = $null
)

$callerLocation = Get-Location
Set-Location -Path $PSScriptRoot

$modulePath = Join-Path $PSScriptRoot 'Modules'
if (-not ($env:PSModulePath -split ';' | Where-Object { $_ -eq $modulePath })) {
    $env:PSModulePath = "$modulePath;$env:PSModulePath"
}
Import-Module -Name MessagingModule -Prefix 'm::' -ErrorAction Stop

function ExitWithError {
    param([string]$Message)
    m::MessageError $Message
    exit 1
}

function Test-LocalBranch([string]$Branch) {
    git show-ref --verify --quiet "refs/heads/$Branch"
    return $LASTEXITCODE -eq 0
}

function Test-RemoteBranch([string]$Branch) {
    git show-ref --verify --quiet "refs/remotes/origin/$Branch"
    return $LASTEXITCODE -eq 0
}

function New-DefaultPullRequestTitle {
    param(
        [Parameter(Mandatory)] [string]$BaseBranch,
        [Parameter(Mandatory)] [string]$HeadBranch
    )

    # Three dots compare the changes introduced by the source branch since its
    # common ancestor with the target branch.  It is the same useful view a PR
    # normally presents to the reviewer.
    $changedFiles = @(git diff --name-only "origin/$BaseBranch...origin/$HeadBranch")
    if ($LASTEXITCODE -ne 0) {
        ExitWithError 'Failed to get the list of changed files for the Pull Request title.'
    }

    if ($changedFiles.Count -lt 5) {
        if ($changedFiles.Count -eq 0) {
            return 'No file changes'
        }
        return "Changed: $($changedFiles -join ', ')"
    }

    return "Changed files: $($changedFiles.Count)"
}

try {
    $workDir = if ($BatDir) {
        try { (Resolve-Path -LiteralPath $BatDir.Trim().Trim('"')).Path }
        catch { ExitWithError "Incorrect path BatDir: '$BatDir'" }
    } else {
        $callerLocation.Path
    }

    $gitRoot = (& git -C $workDir rev-parse --show-toplevel 2>$null).Trim()
    if (-not $gitRoot) { ExitWithError "Not a Git repository at: $workDir" }

    Push-Location $gitRoot
    try {
        $gh = (Get-Command gh -ErrorAction SilentlyContinue).Source
        if (-not $gh) { ExitWithError "GitHub CLI (gh) was not found. Install it and run 'gh auth login'." }
        & $gh auth status 2>$null | Out-Null
        if ($LASTEXITCODE -ne 0) { ExitWithError "GitHub CLI is not authenticated. Run 'gh auth login'." }

        if ([string]::IsNullOrWhiteSpace($Head)) { $Head = Read-Host 'Enter Source Branch (head)' }
        if ([string]::IsNullOrWhiteSpace($Base)) { $Base = Read-Host 'Enter Target Branch (base)' }
        if ([string]::IsNullOrWhiteSpace($Title)) { $Title = Read-Host 'Enter Pull Request Title (empty = automatic)' }
        if ($null -eq $NewBranch) { $NewBranch = Read-Host 'Enter branch to recreate from target HEAD (empty = source branch)' }
        $Head = $Head.Trim(); $Base = $Base.Trim(); $Title = $Title.Trim()
        $NewBranch = if ([string]::IsNullOrWhiteSpace($NewBranch)) { $Head } else { $NewBranch.Trim() }

        if (-not $Head -or -not $Base) { ExitWithError 'Head and base are required.' }
        if ($Head -ieq $Base) { ExitWithError 'Head and base must be different branches.' }
        if (@(git status --porcelain).Count -gt 0) { ExitWithError 'Working tree is not clean. Commit or stash local changes first.' }

        m::MessageAction 'Fetching remote branches...'
        git fetch origin --prune
        if ($LASTEXITCODE -ne 0) { ExitWithError 'Failed to fetch remote branches.' }
        if (-not (Test-RemoteBranch $Base)) { ExitWithError "Target branch 'origin/$Base' does not exist." }

        if (-not (Test-LocalBranch $Head)) {
            if (-not (Test-RemoteBranch $Head)) { ExitWithError "Source branch '$Head' does not exist locally or on origin." }
            git branch --track $Head "origin/$Head"
            if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to create local branch '$Head'." }
        }

        m::MessageAction "Publishing '$Head' to origin..."
        git push origin "$Head`:$Head"
        if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to push '$Head' to origin." }
        git fetch origin --prune
        if ($LASTEXITCODE -ne 0) { ExitWithError 'Failed to refresh remote branches.' }

        $commitCount = [int](git rev-list --count "origin/$Base..origin/$Head")
        if ($LASTEXITCODE -ne 0) { ExitWithError 'Failed to compare source and target branches.' }
        if ($commitCount -gt 0) {
            if ([string]::IsNullOrWhiteSpace($Title)) {
                $Title = New-DefaultPullRequestTitle -BaseBranch $Base -HeadBranch $Head
                m::Message "Automatic Pull Request title: $Title"
            }

            $prNumber = (& $gh pr list --state open --base $Base --head $Head --json number --jq '.[0].number').Trim()
            if (-not $prNumber) {
                m::MessageAction "Creating Pull Request '$Head' -> '$Base'..."
                & $gh pr create --title $Title --body 'Auto-generated PR from AutoPrMerge workflow.' --base $Base --head $Head
                if ($LASTEXITCODE -ne 0) { ExitWithError 'Failed to create Pull Request.' }
                $prNumber = (& $gh pr list --state open --base $Base --head $Head --json number --jq '.[0].number').Trim()
            }
            if (-not $prNumber) { ExitWithError 'Unable to resolve Pull Request number.' }
            m::MessageAction "Merging Pull Request #$prNumber..."
            & $gh pr merge $prNumber --merge --subject "Merge PR #$prNumber $Title"
            if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to merge Pull Request #$prNumber." }
        } else {
            m::Message "No commits to merge from '$Head' into '$Base'. Pull Request is not required."
        }

        git fetch origin --prune
        if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to refresh 'origin/$Base' after merge." }
        if (Test-RemoteBranch $Head) {
            git push origin --delete $Head
            if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to delete old remote branch '$Head'." }
        }
        git fetch origin --prune
        git switch --detach "origin/$Base"
        if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to switch to 'origin/$Base'." }
        if (Test-LocalBranch $NewBranch) { git branch -f $NewBranch "origin/$Base" } else { git branch $NewBranch "origin/$Base" }
        if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to recreate '$NewBranch'." }
        git switch $NewBranch
        if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to switch to '$NewBranch'." }
        if (($NewBranch -ine $Head) -and (Test-LocalBranch $Head)) { git branch -D $Head }
        git push -u origin "$NewBranch`:$NewBranch"
        if ($LASTEXITCODE -ne 0) { ExitWithError "Failed to publish '$NewBranch'." }
        m::Message "Done. 'origin/$Base' and 'origin/$NewBranch' now point to the same commit."
    } finally {
        Pop-Location
    }
} finally {
    Set-Location $callerLocation
}
