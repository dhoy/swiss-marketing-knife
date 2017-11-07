import pysftp

#object
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None
sftp = pysftp.Connection('sftp.crosscountrycomputer.com', username='ccftpEsm', password='STia6#Rx', cnopts=cnopts)
tst = sftp.listdir()
sftp.close()
print(tst)

#Context
# cnopts = pysftp.CnOpts()
# cnopts.hostkeys = None
# with pysftp.Connection('sftp.crosscountrycomputer.com', username='ccftpEsm', password='STia6#Rx', cnopts=cnopts) as sftp:
#     tst = sftp.listdir()
#     print(tst)
#     