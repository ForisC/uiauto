import subprocess
import time

import pytest
import psutil


class Utils:
    @staticmethod
    def launch_notepad():
        print("===== restart notepad ========")
        subprocess.run("taskkill /F /IM notepad.exe", capture_output=True)
        subprocess.Popen("notepad.exe")
        time.sleep(1)

    @staticmethod
    def kill_notepad():
        subprocess.run("taskkill /F /IM notepad.exe", capture_output=True)

    @staticmethod
    def get_image_pids(image_name):
        pids = []
        for proc in psutil.process_iter():
            try:
                if proc.name().lower() == image_name.lower():
                    pids.append(proc.pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return pids


@pytest.fixture(scope="package")
def test_utils():
    return Utils
