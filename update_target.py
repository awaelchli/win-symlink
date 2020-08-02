import subprocess
import glob
import os
from pathlib import Path


def run_command(cmd):
    p = subprocess.Popen(
        ["powershell.exe", cmd],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    out, err, = p.communicate()
    print(err.decode('utf-8'))
    return out.decode('utf-8').strip()


def substitute_drive_letter(link, old, new):
    command = rf'''
    $WshShell = New-Object -comObject WScript.Shell;
    $path = $WshShell.CreateShortcut("{link}").TargetPath;
    $path;
    '''
    file_path = run_command(command)
    file_path = file_path.replace(old, new)
    command = rf'''
    $WshShell = New-Object -comObject WScript.Shell;
    $path = $WshShell.CreateShortcut("{link}");
    $path.TargetPath = "{file_path}";
    $path.Save();
    '''
    run_command(command)


def main():
    path = Path(r"C:\Users\USER\PATH\TO\FOLDER").resolve()
    old = r"D:\SOME\PATH\TO\BE\SUBSTITUTED"
    new = "N:\THE\NEW\PATH"
    os.chdir(path)
    for link in glob.glob("*.lnk"):
        print(link)
        substitute_drive_letter(link, old, new)


if __name__ == '__main__':
    main()
