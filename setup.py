from distutils.text_file import TextFile
from setuptools import find_packages, setup
from uiauto import VERSION

setup(
    name="uiauto-patch",
    version=VERSION,
    url="https://github.com/foris323/uiauto_patch",
    description='Enhanced uiautomation module',
    author='ForisCai',
    author_email='foris323@gmail.com',
    packages=find_packages(),
    setup_requires=["wheel"],
    install_requires=TextFile("requirements.txt").readlines(),
)
