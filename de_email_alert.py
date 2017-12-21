#####################################################################
# Digital Edge Informity                                            #
# httpd://www.digitaledge.net                                       #
#                                                                   #
# This script may not be published or re-used without permission.   #
# Last updated 12/04/2015                                           #
# Email support@digitaledge.net for immediate assistance            #
#####################################################################

import sys
from de_utils import send_email

SMTPSERVER='mc1.digitaledge.net'
SENDFROM='esimacenco@digitaledge.net'
SENDTO=['esimacenco@digitaledge.net','custommonitoring@digitaledge.net']

try:
    send_email(SMTPSERVER, SENDFROM, SENDTO, str(sys.argv[1]), str(sys.argv[2]))
except ValueError, e:
    print e
