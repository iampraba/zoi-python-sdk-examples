# This is a sample Python script.
import os

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from zohosdk.src.com.zoho.exception.sdk_exception import SDKException
from zohosdk.src.com.zoho.user_signature import UserSignature
from zohosdk.src.com.zoho.dc.data_center import DataCenter
from zohosdk.src.com.zoho.api.authenticator.api_key import APIKey
from zohosdk.src.com.zoho.util.constants import Constants
from zohosdk.src.com.zoho.api.logger import Logger
from zohosdk.src.com.zoho import Initializer

from zohosdk.src.com.zoho.officeintegrator.v1 import CompareDocumentParameters, CompareDocumentResponse, \
    InvalidConfigurationException
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations
from zohosdk.src.com.zoho.util import StreamWrapper


class CompareDocument:

    @staticmethod
    def execute():
        CompareDocument.init_sdk()
        compareDocumentParameters = CompareDocumentParameters();

        compareDocumentParameters.set_url1('https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx')
        compareDocumentParameters.set_url2('https://demo.office-integrator.com/zdocs/MS_Word_Document_v1.docx')

        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

        file1Name = "MS_Word_Document_v0.docx"
        # file1Path = ROOT_DIR + "/sample_documents/" + file1Name
        # print('source file1 path : ' + file1Path)
        # compareDocumentParameters.set_document1(StreamWrapper(file_path=file1Path))

        file2Name = "MS_Word_Document_v1.docx"
        # file2Path = ROOT_DIR + "/sample_documents/" + file2Name
        # print('source file1 path : ' + file2Path)
        # compareDocumentParameters.set_document2(StreamWrapper(file_path=file2Path))

        compareDocumentParameters.set_lang("en")
        compareDocumentParameters.set_title(file1Name + " vs " + file2Name)

        v1Operations = V1Operations()
        response = v1Operations.compare_document(compareDocumentParameters)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CompareDocumentResponse):
                    print('Document Compare Session URL  : ' + str(responseObject.get_compare_url()))
                    print('Document Compare Session Delete URL : ' + str(responseObject.get_session_delete_url()))
            elif isinstance(responseObject, InvalidConfigurationException):
                print('Invalid configuration exception.')
                print('Error Code  : ' + str(responseObject.get_code()))
                print("Error Message : " + str(responseObject.get_message()))
                if responseObject.get_parameter_name() is not None:
                    print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                if responseObject.get_key_name() is not None:
                    print("Error Key Name : " + str(responseObject.get_key_name()))
            else:
                print('Compare Document Request Failed')

    @staticmethod
    def init_sdk():
        try:
            # Replace email address associated with your apikey below
            user = UserSignature("john@zylker.com")
            # Update the api domain based on in which data center user register your apikey
            # To know more - https://www.zoho.com/officeintegrator/api/v1/getting-started.html
            environment = DataCenter.Environment("https://api.office-integrator.com", None, None, None)
            # User your apikey that you have in office integrator dashboard
            apikey = APIKey("2ae438cf864488657cc9754a27daa480", Constants.PARAMS)
            # Configure a proper file path to write the sdk logs
            logger = Logger.get_instance(Logger.Levels.INFO, "./logs.txt")

            Initializer.initialize(user, environment, apikey, None, None, logger, None)

        except SDKException as ex:
            print(ex.code)


CompareDocument.execute()
