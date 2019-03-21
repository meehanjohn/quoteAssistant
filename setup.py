try:
    from setuptools import setup, find_packages
except ImportError:
    from distutils.core import setup

with open("requirements.txt",'r') as f:
        requirements = f.read().splitlines()

config = {
    'description': 'Quote Assistant',
    'author': 'John Meehan',
    'url': 'URL to get it at.',
    'download_url': 'Where to download it.',
    'author_email': 'jmeehan@amuneal.com',
    'version': '0.1',
    'install_requires': requirements,
    'packages' : find_packages(),
    'scripts': [],
    'name': 'Quote Assistant'
}

setup(**config)
