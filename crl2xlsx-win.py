#!/usr/bin/python3
from cryptography import x509
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from OpenSSL import crypto
import argparse, sys, tempfile, xlsxwriter, hashlib, datetime, os


def main(argv):
    crl_file = sys.argv[1]
    xlsx_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    
    # Create a new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(xlsx_file)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    # Expand the columns so that the data is visible.
    worksheet.set_column('A:A', 39)
    worksheet.set_column('B:B', 36)
    worksheet.set_column('C:C', 10)
    worksheet.set_column('D:D', 13)
    worksheet.set_column('E:E', 19)
    # Write the column headers.
    worksheet.write('A1', 'Serial (int)', bold)
    worksheet.write('B1', 'Serial (hex)', bold)
    worksheet.write('C1', 'Date', bold)
    worksheet.write('D1', 'Time (UTC)', bold)
    worksheet.write('E1', 'Reason', bold)
    # Start from first row after headers.
    row = 1
    date_format = workbook.add_format({'num_format': 'dd-mm-yyyy', 'align': 'left'})
    time_format = workbook.add_format({'num_format': 'hh:mm:ss', 'align': 'left'})
    text_format = workbook.add_format({'num_format': '@', 'align': 'left'})
    
    with open(crl_file, "rb") as in_file:
        crl_obj = crypto.load_crl(crypto.FILETYPE_ASN1, in_file.read())
        crl_contents = crypto.dump_crl(crypto.FILETYPE_PEM, crl_obj)

        crl = x509.load_pem_x509_crl(crl_contents, default_backend())

        for revoked_cert in crl:
            serial_int = revoked_cert.serial_number
            serial_hex = '{:x}'.format(serial_int)
            revocation_datetime = revoked_cert.revocation_date

            try:
                reason_ext = revoked_cert.extensions.get_extension_for_oid(x509.CRLEntryExtensionOID.CRL_REASON)
            except x509.extensions.ExtensionNotFound:
                reason = ""
            else:
                reason = reason_ext.value.reason.value

            # Add revocation to worksheet
            worksheet.write_string(row, 0, str(serial_int), text_format)
            worksheet.write_string(row, 1, serial_hex, text_format)
            worksheet.write_datetime(row, 2, revocation_datetime, date_format)
            worksheet.write_datetime(row, 3, revocation_datetime, time_format)
            worksheet.write_string(row, 4, reason, text_format)
            row += 1

        worksheet.autofilter('A1:E' + str(row-1))
        # Close Excel workbook
        workbook.close()
        os.system('start excel.exe ' + xlsx_file.name)

if __name__ == "__main__":
    main(sys.argv[1:])