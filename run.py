import sys
from os import execv
from pathlib import Path
from subprocess import run

from text_cleanup import cli


def _is_supported_py() -> bool:
    major, minor = sys.version_info[:2]
    return major == 3 and 11 <= minor <= 13


def _venv311_python() -> Path:
    return Path(__file__).resolve().parent / ".venv311" / "Scripts" / "python.exe"


def _relaunch_with_venv311() -> bool:
    venv_py = _venv311_python()
    if not venv_py.exists():
        return False
 
 
    current_py = Path(sys.executable).resolve()
    if venv_py.resolve() == current_py:
        return False

    print("检测到 .venv311，自动切换解释器启动...")
    try:
        execv(str(venv_py), [str(venv_py), *sys.argv])
    except Exception:
        run([str(venv_py), *sys.argv], check=False)
        return True
    return True


def _install_missing_gui_deps() -> bool:
    requirements_file = Path(__file__).resolve().parent / "requirements-py311.txt"
    if requirements_file.exists():
        command = [sys.executable, "-m", "pip", "install", "-r", str(requirements_file)]
    else:
        command = [sys.executable, "-m", "pip", "install", "customtkinter", "python-docx", "reportlab"]

    print("正在为当前解释器自动安装 GUI 依赖，请稍候...")
    result = run(command, check=False)
    return result.returncode == 0


def _normalize_cli_args() -> None:
    if "--text" not in sys.argv:
        return

    index = sys.argv.index("--text")
    if index + 1 >= len(sys.argv):
        return

    next_arg = sys.argv[index + 1]
    if next_arg.startswith("--"):
        return

    tail = []
    for item in sys.argv[index + 2 :]:
        if item.startswith("--"):
            break
        tail.append(item)

    if not tail:                   
        return

    merged = " ".join([next_arg, *tail])
    new_argv = sys.argv[: index + 1] + [merged]
    consumed = index + 2 + len(tail)
    new_argv.extend(sys.argv[consumed:])
    sys.argv = new_argv


def main() -> None:
    if not _is_supported_py():
        print("当前 Python 版本不推荐。请使用 Python 3.11（推荐）或 3.12/3.13。")
        print("建议在项目目录创建并使用 .venv311，再安装 requirements-py311.txt")

        if _relaunch_with_venv311():
            return

        if len(sys.argv) > 1 and sys.argv[1] == "cli":
            return

    if len(sys.argv) > 1 and sys.argv[1] == "cli":
        _normalize_cli_args()
        sys.argv = [sys.argv[0], *sys.argv[2:]]
        cli()
        return

    try:
        from app_gui import run_gui
    except ModuleNotFoundError as error:
        missing = str(error)
        if "customtkinter" in missing:
            if _relaunch_with_venv311():
                return

            if _install_missing_gui_deps():
                try:
                    from app_gui import run_gui
                except ModuleNotFoundError:
                    print(f"缺少 customtkinter，请先安装：{sys.executable} -m pip install customtkinter")
                    return
                run_gui()
                return

            print(f"缺少 customtkinter，请先安装：{sys.executable} -m pip install customtkinter")
            return
        raise

    run_gui()


if __name__ == "__main__":
    main()
