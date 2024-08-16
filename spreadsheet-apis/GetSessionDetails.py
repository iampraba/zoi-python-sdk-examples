from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import SessionInfo, \
    SessionMeta, InvalidConfigurationException, CreateSheetParameters, \
    CreateSheetResponse, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class GetSessionDetails:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-sheet-session-information.html
    @staticmethod
    def execute():
        GetSessionDetails.init_sdk()
        createSpreadsheetParams = CreateSheetParameters()

        print('Creating a spreadsheet to demonstrate get spreadsheet session information api')
        v1Operations = V1Operations()
        response = v1Operations.create_sheet(createSpreadsheetParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateSheetResponse):
                    sessionId = str(responseObject.get_session_id())
                    print('Created Spreadsheet Session ID : ' + sessionId)

                    sessionInfoResponse = v1Operations.get_sheet_session(sessionId)

                    if sessionInfoResponse is not None:
                        print('Status Code: ' + str(sessionInfoResponse.get_status_code()))
                        sessionInfoObj = sessionInfoResponse.get_object()

                        if sessionInfoObj is not None:
                            if isinstance(sessionInfoObj, SessionMeta):
                                print('Session Status : ' + str(sessionInfoObj.get_status()))

                                sessionInfo = sessionInfoObj.get_info()

                                if isinstance(sessionInfo, SessionInfo):
                                    print('Spreadsheet ID : ' + str(sessionInfo.get_document_id()))
                                    print('Spreadsheet Session Created Time : ' + str(sessionInfo.get_created_time()))
                                    print('Spreadsheet Session Created Timestamp : ' + str(sessionInfo.get_created_time_ms()))
                                    print('Spreadsheet Session Expiry Time : ' + str(sessionInfo.get_expires_on()))
                                    print('Spreadsheet Session Expiry Timestamp : ' + str(sessionInfo.get_expires_on_ms()))
                            elif isinstance(sessionInfoObj, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(sessionInfoObj.get_code()))
                                print("Error Message : " + str(sessionInfoObj.get_message()))
                                if sessionInfoObj.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(sessionInfoObj.get_parameter_name()))
                                if sessionInfoObj.get_key_name() is not None:
                                    print("Error Key Name : " + str(sessionInfoObj.get_key_name()))
                            else:
                                print('Get Spreadsheet Session Details Request Failed')
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Create Spreadsheet Request Failed')

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


GetSessionDetails.execute()
