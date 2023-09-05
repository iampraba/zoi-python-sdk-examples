# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from zohosdk.src.com.zoho.exception.sdk_exception import SDKException
from zohosdk.src.com.zoho.user_signature import UserSignature
from zohosdk.src.com.zoho.dc.data_center import DataCenter
from zohosdk.src.com.zoho.api.authenticator.api_key import APIKey
from zohosdk.src.com.zoho.util.constants import Constants
from zohosdk.src.com.zoho.api.logger import Logger
from zohosdk.src.com.zoho import Initializer

from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentInfo, UserInfo, CallbackSettings
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

import time
import os

from zohosdk.src.com.zoho.util import StreamWrapper


class EditDocument:

    @staticmethod
    def execute():
        EditDocument.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        # Add document meta to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("Existing Document")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        createDocumentParams.set_document_info(documentInfo)

        # Add User meta to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1000")
        userInfo.set_display_name("Prabakaran R")

        createDocumentParams.set_user_info(userInfo)

        callbackSettings = CallbackSettings()
        saveUrlParams = {}

        saveUrlParams['auth_token'] = '1234'
        saveUrlParams['id'] = '123131'

        callbackSettings.set_save_url_params(saveUrlParams)

        saveUrlHeaders = {}

        saveUrlHeaders['access_token'] = '12dweds32r42wwds34'
        saveUrlHeaders['client_id'] = '12313111'

        callbackSettings.set_save_url_headers(saveUrlHeaders)

        callbackSettings.set_retries(1)
        callbackSettings.set_timeout(10000)
        callbackSettings.set_save_format("zdoc")
        callbackSettings.set_http_method_type("post")
        callbackSettings.set_save_url(
            "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286")

        createDocumentParams.set_callback_settings(callbackSettings)

        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        filePath = ROOT_DIR + "/sample_documents/Graphic-Design-Proposal.docx"

        print('Path for source file to be edited : ' + filePath)

        createDocumentParams.set_document(StreamWrapper(file_path=filePath))

        # createDocumentParams.set_url('https://demo.office-integrator.com/zdocs/LabReport.zdoc')

        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            response_object = response.get_object()

            if response_object is not None:
                if isinstance(response_object, CreateDocumentResponse):
                    print('Document Id : ' + str(response_object.get_document_id()))
                    print('Document URL : ' + str(response_object.get_document_url()))

    @staticmethod
    def init_sdk():
        try:
            user = UserSignature("john@zylker.com")
            environment = DataCenter.Environment("https://api.office-integrator.com", None, None, None)
            apikey = APIKey("2ae438cf864488657cc9754a27daa480", Constants.PARAMS)
            logger = Logger.get_instance(Logger.Levels.INFO, "./logs.txt")

            Initializer.initialize(user, environment, apikey, None, None, logger, None)

        except SDKException as ex:
            print(ex.code)


EditDocument.execute()
