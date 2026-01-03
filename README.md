#ACLauncher

ACLauncher is a dual-configuration launcher for ACLog to provide two unique environments and log files installed on a single computer.

## Installing ps2exe

ps2exe compiles the PowerShell script into an executable file. From an administrator window of PowerShell, install the ps2exe module.

```install-module ps2exe```

## Compiling the ACLauncher

Once ps2exe is installed, navigate to the directory with the script in it. Invoke the compiler with the following outputs.
```
Invoke-PS2EXE -inputFile 'ACLauncher.ps1' -outputFile 'ACLauncher.exe' -iconFile 'aclauncher.ico' -title 'ACLauncher' -description 'Launches his and hers ACLog' -company 'Custom Digital Services' -product 'ACLauncher' -version '1.1.0' -verbose -noConsole
```
