from distutils.text_file import TextFile
from setuptools import find_packages, setup
from uiauto import VERSION

setup(
    name="uiauto-patch",
    version=VERSION,
    description='Enhanced uiautomation module',
    license='Apache 2.0',
    author='ForisCai',
    author_email='foris323@gmail.com',
    keywords="windows ui automation uiautomation inspect",
    url="https://github.com/foris323/uiauto_patch",
    platforms='Windows Only',
    packages=find_packages(),
    long_description='Enhanced python UIAutomation for Windows. Supports Python3.4+, x86, x64',
    setup_requires=["wheel"],
    classifiers=[
        "Programming Language :: Python :: 3.4",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
    ],
    install_requires=TextFile("requirements.txt").readlines(),
)
