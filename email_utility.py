import os
import win32com.client as win32
import re
import sys
import shutil
from os import listdir
from win32com.client import constants as c


def outlook():
    try:
        app = win32.gencache.EnsureDispatch("Outlook.Application")
        return app
    except AttributeError:
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        xl = client.gencache.EnsureDispatch('Excel.Application')


def replace_spl_characters(str_):
    """

    :param str_:
    :return:
    """
    return str_.replace('/', '').replace(':', '')


def create_email(send_from, subject, body, params):
    curr_dir = os.getcwd()
    doc_names = [replace_spl_characters(str_=name) for name in params.keys()]
    dict_ = {replace_spl_characters(str_=name): value for name, value in params.items()}
    for file_name in listdir(curr_dir):
        user_name = file_name.split('.')[0].split('_')[-1]
        if user_name in doc_names:
            app = outlook()
            email = app.CreateItem(c.olMailItem)
            email.SentOnBehalfOfName = send_from
            file_attach = os.path.join(curr_dir, file_name)
            email.Attachments.Add(file_attach, c.olByValue)
            email.Subject = subject
            email.Body = "Hello Everyone"
            email.HTMLBody = body
            email.To = dict_[user_name]
            email.Close(c.olSave)


if __name__ == '__main__':
    create_email(
        "narasimha.bandaru@maverickcap.com",
        "foo@bar.com;baz@bar.com",
        "Test",
        "Body",
        r"c:\temp\2020.txt"
    )