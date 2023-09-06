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

from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentDeleteSuccessResponse, AllSessionsResponse, SessionInfo, \
    SessionMeta, SessionUserInfo, DocumentMeta, InvalidConfigurationException
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations


class GetDocumentDetails:

    @staticmethod
    def execute():
        GetDocumentDetails.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        print('Creating a document to demonstrate get all document information api')
        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    documentId = str(responseObject.get_document_id())
                    print('Created Document ID : ' + documentId)

                    documentInfoResponse = v1Operations.get_document_info(documentId)

                    if documentInfoResponse is not None:
                        print('Status Code: ' + str(documentInfoResponse.get_status_code()))
                        documentInfoObj = documentInfoResponse.get_object()

                        if documentInfoObj is not None:
                            if isinstance(documentInfoObj, DocumentMeta):
                                print('Document ID : ' + str(documentInfoObj.get_document_id()))
                                print('Document Name : ' + str(documentInfoObj.get_document_name()))
                                print('Document Type : ' + str(documentInfoObj.get_document_type()))
                                print('Document Collaborators Count : ' + str(documentInfoObj.get_collaborators_count()))
                                print('Document Created Time : ' + str(documentInfoObj.get_created_time()))
                                print('Document Created Timestamp : ' + str(documentInfoObj.get_created_time_ms()))
                                print('Document Expiry Time : ' + str(documentInfoObj.get_expires_on()))
                                print('Document Expiry Timestamp : ' + str(documentInfoObj.get_expires_on_ms()))
                            elif isinstance(documentInfoObj, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(documentInfoObj.get_code()))
                                print("Error Message : " + str(documentInfoObj.get_message()))
                                if documentInfoObj.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(documentInfoObj.get_parameter_name()))
                                if documentInfoObj.get_key_name() is not None:
                                    print("Error Key Name : " + str(documentInfoObj.get_key_name()))
                            else:
                                print('Get Document Details API Request Failed')
            elif isinstance(responseObject, InvalidConfigurationException):
                print('Invalid configuration exception.')
                print('Error Code  : ' + str(responseObject.get_code()))
                print("Error Message : " + str(responseObject.get_message()))
                if responseObject.get_parameter_name() is not None:
                    print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                if responseObject.get_key_name() is not None:
                    print("Error Key Name : " + str(responseObject.get_key_name()))
            else:
                print('Create Document Request Failed')

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


GetDocumentDetails.execute()
