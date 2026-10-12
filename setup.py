import pathlib
from setuptools import find_packages, setup


# The directory containing this file
HERE = pathlib.Path(__file__).parent

# The text of the README file
README = (HERE / "README.md").read_text(encoding="utf-8")

setup(
    name="sendemail",
    version="0.0.2",
    python_requires=">=3.8",
    packages=find_packages(),
    py_modules=["send_outlookemail", "mail_merge", "outlook_gui"],
    license="Private",
    description="send email via outlook",
    long_description=README,
    long_description_content_type="text/markdown",
    author="sukhbinder",
    author_email="sukh2010@yahoo.com",
    keywords=["send", "windows", "email", "outlook"],
    entry_points={
        "console_scripts": [
            "sendmail = send_outlookemail:main",
        ],
        "gui_scripts": ["sendmail-gui = outlook_gui:main"],
    },
    install_requires=["pywin32; sys_platform == 'win32'"],
    extras_require={
        "test": [
            "pytest",]
    },
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
    ],
)
