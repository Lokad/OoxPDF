
## .NET Commands

Use `--tl:off` to avoid dynamic terminal logger output.

```powershell
dotnet restore Foo.slnx --tl:off -v minimal
dotnet build   Foo.slnx --tl:off --nologo -v minimal
dotnet test    tests/Foo.Tests/Foo.Tests.csproj --tl:off --nologo -v minimal
dotnet pack    src/Foo/Foo.csproj --tl:off --nologo -v minimal --no-restore
```

## ExecPlans

When writing complex features or significant refactors, use an ExecPlan (as described in PLANS.md) from design to implementation.

## Autonomy Policy

If you're working towards goals, do NOT end your turn. This allows for continuous autonomous work.

The user will interrupt you when required, but they will mostly provide steering messages.

Do not pester the user by ending your turn after a unit of work, as that requires them to keep nudging you to keep working.

You MUST continue working autonomously towards any known objectives until the user interrupts you. Do NOT end your turn until there is absolutely nothing left to do.