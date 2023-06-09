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

from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentInfo
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

import time

from zohosdk.src.com.zoho.util import StreamWrapper


class CreateDocument:

    @staticmethod
    def execute():
        CreateDocument.init_sdk()
        createDocumentParams = CreateDocumentParameters()
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("Untilted Document")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())
        createDocumentParams.set_document_info(documentInfo)

        createDocumentParams.set_document(StreamWrapper(file_path="/Users/praba-2086/Desktop/writer.docx"))
        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            response_object = response.get_object()

            if response_object is not None:

                if isinstance(response_object, CreateDocumentResponse):
                    print('Docment Id : ' + str(response_object.get_document_id()))
                    print('Docment URL : ' + str(response_object.get_document_url()))

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


CreateDocument.execute()
