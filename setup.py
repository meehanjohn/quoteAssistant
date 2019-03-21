try:
    from setuptools import setup, find_packages
except ImportError:
    from distutils.core import setup

with open("requirements.txt",'r') as f:
        requirements = f.read().splitlines()

config = {
    'description': 'Quote Assistant',
    'author': 'John Meehan',
    'url': 'https://github.com/meehanjohn/quoteAssistant',
    'download_url': 'https://github.com/meehanjohn/quoteAssistant.git',
    'author_email': 'jmeehan@amuneal.com',
    'version': '0.1',
    'install_requires': requirements,
    'packages' : find_packages(),
    'scripts': [],
    'name': 'Quote Assistant'
}

setup(**config)
