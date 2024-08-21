from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer

from officeintegrator.src.com.zoho.officeintegrator.v1 import DocumentSessionDeleteSuccessResponse, \
    InvalidConfigurationException, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class DeleteDocumentSession:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-delete-user-session.html
    @staticmethod
    def execute():
        DeleteDocumentSession.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        print('Creating a document to demonstrate document session delete api')
        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    documentSessionId = str(responseObject.get_session_id())
                    print('Document Session ID to be deleted : ' + documentSessionId)

                    deleteApiResponse = v1Operations.delete_session(documentSessionId)

                    if deleteApiResponse is not None:
                        print('Status Code: ' + str(deleteApiResponse.get_status_code()))
                        deleteResponseObject = deleteApiResponse.get_object()

                        if deleteResponseObject is not None:
                            if isinstance(deleteResponseObject, DocumentSessionDeleteSuccessResponse):
                                print('Document Session delete status : ' + str(deleteResponseObject.get_session_deleted()))
                            elif isinstance(deleteResponseObject, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(deleteResponseObject.get_code()))
                                print("Error Message : " + str(deleteResponseObject.get_message()))
                                if deleteResponseObject.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(deleteResponseObject.get_parameter_name()))
                                if deleteResponseObject.get_key_name() is not None:
                                    print("Error Key Name : " + str(deleteResponseObject.get_key_name()))
                            else:
                                print('Delete Document Session Request Failed')
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Document Creation Request Failed')

    @staticmethod
    def init_sdk():
        try:
            #Sdk application log configuration
            logger = Logger.get_instance(Logger.Levels.INFO, "./logs.txt")
            #Update this apikey with your own apikey signed up in office integrator service
            auth = Auth.Builder().add_param("apikey", "2ae438cf864488657cc9754a27daa480").set_authentication_schema(Authentication.TokenFlow()).build()
            tokens = [ auth ]
            # Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html
            environment = APIServer.Production("https://api.office-integrator.com")

            Initializer.initialize(environment, tokens,None, None, logger, None)
        except SDKException as ex:
            print(ex.code)

DeleteDocumentSession.execute()
