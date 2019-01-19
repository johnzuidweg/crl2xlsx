# crl2xlsx
Convert a DER-encoded certificate revocation list (CRL) to an XLSX-file.

# crl2xlsx.py
Python3 implementation using cryptography, pyOpenSSL and XlsxWriter.
Creates an .xlsx file listing CRL contents

```
$ pip install cryptography, pyOpenSSL, XlsxWriter
```
```
$ crl2xlsx.py <CRL file (DER)> <name for new .xlsx file>
```

# crl2xlsx-win.py
Windows-targeted Python3 implementation using cryptography, pyOpenSSL and XlsxWriter.
Assuming MS Excel is installed. Creates a temporary .xlsx file listing CRL contents and opens it in Excel.
Download prebuilt binary [crl2xlsx.exe](https://github.com/johnzuidweg/crl2xlsx/releases/download/1/crl2xlsx.exe) (x86)
