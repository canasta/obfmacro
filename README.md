# obfmacro
Detecting obfuscation in VBA macro

We use VBA macro files(.cls and .bas) as target.
So we have to extract macro files from document files before detection.
We use olevba(https://github.com/decalage2/oletools/wiki/olevba) library to extract.
You can find python code for extracting macro in "extract_vba.py"
