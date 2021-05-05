from setuptools import find_packages, setup

setup(
    name="uiauto-patch",
    version="0.0.1",
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
    install_requires=[
        "uiautomation==2.0.11",
        "psutil",
        "pywin32",
    ],
)
