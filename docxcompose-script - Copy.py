#!"c:\users\lathan larue\appdata\local\programs\python\python37-32\python.exe"
# EASY-INSTALL-ENTRY-SCRIPT: 'docxcompose==1.0.0a16','console_scripts','docxcompose'
__requires__ = 'docxcompose==1.0.0a16'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('docxcompose==1.0.0a16', 'console_scripts', 'docxcompose')()
    )
