from setuptools import setup

setup(
    name='shouyu',
    version='0.2',
    description='Quickly record your mind to MS/WPS Excel file by hot keys, building your personal knowledge base.',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/pmh905001/shouyu',
    project_urls={'Bug Tracker': 'https://github.com/pmh905001/shouyu/issues'},
    author='Peng Ming Hua',
    author_email='pmh905001@126.com',
    license='MIT',
    classifiers=[
        'Programming Language :: Python :: 3.8',
        'License :: OSI Approved :: MIT License',
        "Topic :: Utilities",
        'Operating System :: Microsoft',
        'Operating System :: Microsoft :: Windows :: Windows 10',
        'Operating System :: Microsoft :: Windows :: Windows 11'
    ],
    install_requires=[
        'pyperclip',
        'keyboard',
        'openpyxl',
        'Pillow',
        'psutil',
        'pystray',
        'pytest',
        'pywin32',
        'PyAutoGUI'
    ],
    py_modules=['shouyu']
)

# Publish commands
# https://packaging.python.org/tutorials/packaging-projects/
# pip install --upgrade pip build twine
# python -m build
# python -m twine upload dist/*
