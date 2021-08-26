from os import rename
from pathlib import Path

from pst2eml import eml_get_parameters, rename_eml

def nonreg():
    tests = [{"name":"big5_subject",
            "fp":"D:/gitw/pyPST2EML/test/big5_subject.eml",
            "subj":"[EXTERNAL]AVIS安維斯汽車租賃股份有限公司-電子發票證明聯開立通知_MK02836722"},
            {"name":"sent_date_non_supported_tz",
            "fp":"D:/gitw/pyPST2EML/test/sent_date_non_supported_tz.eml",
            "sent_date":"Fri, 15 Feb 2008 13:51:50"},
            {"name":"cannot_get_senddate",
            "sent_date":"Wednesday, February 02, 2005 2:02 PM"},
            {"name":"1081",
            "subj":"http --www.linkedin.com-share-viewLink=&sid=s799124597&url=http%3A%2F%2Fwww%2Eforbes%2Ecom%2Fsites%2Fhaydnshaughnessy%2F2012%2F01%2F04%2Fwhat-do-social-media-influencers-do-that-you-dont-but-could%2F&urlhash=wTFE&pk=nhome-chron-split-feed-items&pp=&poster",
            "rename":"Y"},
            {"name":"e_acute",
            "subj":"bonne année"},
            {"name":"error_sent_date2",
            "sent_date":"Wed, 9 feb 2005 16:08:18 +0100"},
            {"name":"cannot_get_sent_date2",
            "sent_date":"Wed, 19 Sep 2001 16:42:32"},
            {"name":"senddate3",
            "sent_date":"Fri, 24 Aug 2001 09:53:36 +0100"},
            {"name":"...",
            "rename":"Y"},
            {"name":"subject",
            "subject":"",
            "rename":"Y"},
            {"name":"[1].[NEPTUNE][ASF-DSP] ASF-DSP documents for BL4&BL5",
            "subject":"",
            "rename":"Y"}
            ]

    errors=0
    for test in tests:
        test_name = test["name"]
        if "fp" in test:
            eml_fp = test["fp"]
        else:
            eml_fp = f"D:/gitw/pyPST2EML/test/{test_name}.eml"

        
        send_to, sent_from, subject, sent_date = eml_get_parameters(eml_fp,debug=True)
        if "subj" in test:
            subj = test["subj"]
            try:
                assert subject==subj
            except:
                errors+=1
                print(f"failed test {test_name} - wrong subject decoding")
                print(f"{subject} vs should be {subj}")
        if "sent_date" in test:
            sent = test["sent_date"]
            try:
                assert sent == sent_date
            except:
                errors+=1
                print(f"failed test {test_name} - wrong sent_date")
                raise
        print(44,test)
        if "rename" in test:
            try:
                new_eml_fp = rename_eml(eml_fp,subject,debug=True)
            except:
                raise
            else:
                print("change back",new_eml_fp,eml_fp)
                rename_eml(new_eml_fp,Path(eml_fp).stem,debug=True)


    print(f"finished with {errors} error(s)")
    with open("D:/gitw/pyPST2EML/test/test1.txt",'w',encoding="utf-8") as fo:
        fo.write(subject)

nonreg()