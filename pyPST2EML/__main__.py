""" PYTHON STANDARD LIBRARY """
import argparse
from os.path import abspath, exists, dirname, join
import logging

# own modules
from .pst2eml import make_eml_search_friendly, pst_2_eml, str2bool

folder = abspath(join(dirname(__file__), "test"))

parser = argparse.ArgumentParser(description='Processing emails for easier desktop searches')
parser.add_argument('--pst', dest='PST_nEML',type=str2bool,
              help='processing .PST (Y) or folder of .eml/.ics (N) (default: N)', default = "N")
parser.add_argument('--folder', "-f", dest='folder',type=str , default = "",
              help="folder where processing will occur, if none then test folder used:")
parser.add_argument('--filename',"-n", dest='fn',type=str,action="store", default = "c:/",
              help='file to be processed (required for .PST processing, optional for .eml processing)')
parser.add_argument('--embedded_libpst', dest='embedded_libpst',type=str, default = "Y",
              help='use the package provided libpst (default: Y)')
parser.add_argument("--test","-t",dest='test',type=str, default = "N",
              help='will run over test files')
parser.add_argument("--verbose","-v",dest='verbosity',type=int,default=2)

args,  unknown = parser.parse_known_args()
myargs = vars(args)

if myargs["verbosity"]==4:
  logging.basicConfig(level=logging.DEBUG)
  logging.info("logging set to DEBUG")
  stop_error = True
else:
  logging.basicConfig(level=logging.WARN)
  logging.info("logging set to WARN")

if myargs["test"]=="Y":
    myargs["PST_nEML"]=False
    myargs["folder"]=folder

if myargs["PST_nEML"]:
  pst_2_eml(**myargs)
else:
    if exists(myargs["folder"]):
        make_eml_search_friendly(myargs["folder"])
    else:
        print("need to provide .pst file path or .eml folder for processing")