import subprocess
from app.config import EDGE_EXE, USER_DATA_DIR, EDGE_PROFILE_DIR, WIS_URL

def setup_profile():
    cmd = [
        EDGE_EXE,
        f"--user-data-dir={USER_DATA_DIR}",
        f"--profile-directory={EDGE_PROFILE_DIR}",
        WIS_URL
    ]
    subprocess.Popen(cmd)
    print("Normal Edge opened with your profile.")
    print("Login manually there.")