from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer

from officeintegrator.src.com.zoho.officeintegrator.v1 import AllSessionsResponse, SessionInfo, \
    SessionMeta, SessionUserInfo, InvalidConfigurationException, UserInfo, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from officeintegrator.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class GetAllSessions:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-get-document-sessions.html
    @staticmethod
    def execute():
        GetAllSessions.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1000")
        userInfo.set_display_name("User 1")

        createDocumentParams.set_user_info(userInfo)

        print('Creating a document to demonstrate get all document session information api')
        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    documentId = str(responseObject.get_document_id())
                    print('Document ID to be deleted : ' + documentId)

                    allSessionResponse = v1Operations.get_all_sessions(documentId)

                    if allSessionResponse is not None:
                        print('Status Code: ' + str(allSessionResponse.get_status_code()))
                        allSessionsResponseObj = allSessionResponse.get_object()

                        if allSessionsResponseObj is not None:
                            if isinstance(allSessionsResponseObj, AllSessionsResponse):
                                print('Document ID : ' + str(allSessionsResponseObj.get_document_id()))
                                print('Document Name : ' + str(allSessionsResponseObj.get_document_name()))
                                print('Document Type : ' + str(allSessionsResponseObj.get_document_type()))
                                print('Document Created Time : ' + str(allSessionsResponseObj.get_created_time()))
                                print('Document Created Timestamp : ' + str(allSessionsResponseObj.get_created_time_ms()))
                                print('Document Expiry Time : ' + str(allSessionsResponseObj.get_expires_on()))
                                print('Document Expiry Timestamp : ' + str(allSessionsResponseObj.get_expires_on_ms()))

                                sessions = allSessionsResponseObj.get_sessions()

                                print('\n---- Document Session Details ----\n')

                                for session in sessions:
                                    if isinstance(session, SessionMeta):
                                        print('Session Status : ' + str(session.get_status()))

                                        sessionInfo = session.get_info()

                                        if isinstance(sessionInfo, SessionInfo):
                                            print('Session User ID : ' + str(sessionInfo.get_session_url()))

                                        sessionUserInfo = session.get_user_info()

                                        if isinstance(sessionUserInfo, SessionUserInfo):
                                            print('Session User ID : ' + str(sessionUserInfo.get_user_id()))
                                            print('Session Display Name : ' + str(sessionUserInfo.get_display_name()))
                            elif isinstance(allSessionsResponseObj, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(allSessionsResponseObj.get_code()))
                                print("Error Message : " + str(allSessionsResponseObj.get_message()))
                                if allSessionsResponseObj.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(allSessionsResponseObj.get_parameter_name()))
                                if allSessionsResponseObj.get_key_name() is not None:
                                    print("Error Key Name : " + str(allSessionsResponseObj.get_key_name()))
                            else:
                                print('Create All Session Details Request Failed')
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

GetAllSessions.execute()