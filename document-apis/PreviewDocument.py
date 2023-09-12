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

from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentInfo, UserInfo, CallbackSettings, DocumentDefaults, \
    EditorSettings, UiOptions, InvalidConfigurationException, PreviewParameters, PreviewDocumentInfo, PreviewResponse
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

import time
import os

from zohosdk.src.com.zoho.util import StreamWrapper


class PreviewDocument:

    # Refer Preview API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-preview-document.html
    @staticmethod
    def execute():
        PreviewDocument.init_sdk()
        previewParameter = PreviewParameters()

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = PreviewDocumentInfo()
        documentInfo.set_document_name("New Document")

        previewParameter.set_document_info(documentInfo)

        # Optional Configuration - Configure permission values for session
        # based of you application requirement
        permissions = {}

        permissions["document.print"] = True

        previewParameter.set_permissions(permissions)

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/Graphic-Design-Proposal.docx"
        # print('Path for source file to be edited : ' + filePath)
        # previewParameter.set_document(StreamWrapper(file_path=filePath))

        previewParameter.set_url('https://demo.office-integrator.com/zdocs/LabReport.zdoc')

        v1Operations = V1Operations()
        response = v1Operations.create_document_preview(previewParameter)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, PreviewResponse):
                    print('Document Id : ' + str(responseObject.get_document_id()))
                    print('Document Session ID : ' + str(responseObject.get_session_id()))
                    print('Document Preview URL : ' + str(responseObject.get_preview_url()))
                    print('Document Session Delete URL : ' + str(responseObject.get_session_delete_url()))
                    print('Document Delete URL : ' + str(responseObject.get_document_delete_url()))
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Preview Document Request Failed')

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


PreviewDocument.execute()
