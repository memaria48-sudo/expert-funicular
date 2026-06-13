import os
import subprocess
from datetime import datetime
from pathlib import Path

from prototype_v2_final import main as prototype_main
from sor_site_postprocess import postprocess_site


WEBSITE_COMMIT_PATHS = [
    "index.html",
    "zirp_dashboard.html",
    "sor-logo.png",
    "zirp_berichte/archive.html",
    "zirp_berichte/scip_archive.html",
    "zirp_berichte/signals.html",
    "zirp_berichte/sor-logo-full.png",
    "zirp_berichte/sor-logo.png",
    "zirp_berichte/sor-tab-o.png",
    "zirp_berichte/zirp_dashboard.html",
    "zirp_berichte/zirp_meeting_recommendations_latest.html",
    "zirp_berichte/zirp_meeting_recommendations_ai_latest.html",
]


def website_commit_paths(project_dir: Path) -> list[str]:
    paths = list(WEBSITE_COMMIT_PATHS)
    reports_dir = project_dir / "zirp_berichte"
    if reports_dir.exists():
        paths.extend(
            str(path.relative_to(project_dir)).replace("\\", "/")
            for path in sorted(reports_dir.glob("zirp_meeting_recommendations*.html"))
            if is_publishable_html(path)
        )
    return sorted(dict.fromkeys(paths))


def is_publishable_html(path: Path) -> bool:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return False
    mojibake_markers = ("\u00c3", "\u00c2", "\u00e2", "\u20ac", "\ufffd")
    return not any(marker in text for marker in mojibake_markers)


def main() -> None:
    prototype_main()
    postprocess_site()
    auto_git_commit_website()


def auto_git_commit_website() -> None:
    project_dir = Path(__file__).resolve().parent
    website_paths = website_commit_paths(project_dir)
    commit_message = f"Auto update: website output {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    try:
        status = subprocess.run(
            ["git", "-C", str(project_dir), "status", "--porcelain", "--", *website_paths],
            capture_output=True,
            text=True,
            check=False,
        )
        if status.returncode != 0:
            print("Auto git website commit: skipped (git status failed).")
            return
        if not status.stdout.strip():
            print("Auto git website commit: no website changes.")
            return

        subprocess.run(
            ["git", "-C", str(project_dir), "add", "-f", "--", *website_paths],
            check=True,
        )
        subprocess.run(
            ["git", "-C", str(project_dir), "commit", "-m", commit_message],
            check=True,
        )
        print("Auto git website commit: commit created.")

        push = subprocess.run(
            ["git", "-C", str(project_dir), "push"],
            capture_output=True,
            text=True,
            check=False,
        )
        if push.returncode == 0:
            print("Auto git website commit: push successful.")
        else:
            print("Auto git website commit: push failed, commit stays local.")
    except Exception as exc:
        print(f"Auto git website commit: skipped ({exc})")


if __name__ == "__main__":
    main()
