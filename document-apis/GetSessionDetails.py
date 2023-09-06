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

from zohosdk.src.com.zoho.officeintegrator.v1 import SessionInfo, SessionMeta, SessionUserInfo, \
    InvalidConfigurationException, UserInfo
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class GetSessionDetails:

    @staticmethod
    def execute():
        GetSessionDetails.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        print('Creating a document to demonstrate get document session information api')

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1000")
        userInfo.set_display_name("User 1")

        createDocumentParams.set_user_info(userInfo)

        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    sessionId = str(responseObject.get_session_id())
                    print('Created Document Session ID : ' + sessionId)

                    sessionInfoResponse = v1Operations.get_session(sessionId)

                    if sessionInfoResponse is not None:
                        print('Status Code: ' + str(sessionInfoResponse.get_status_code()))
                        sessionInfoObj = sessionInfoResponse.get_object()

                        if sessionInfoObj is not None:
                            if isinstance(sessionInfoObj, SessionMeta):
                                print('Session Status : ' + str(sessionInfoObj.get_status()))

                                sessionInfo = sessionInfoObj.get_info()

                                if isinstance(sessionInfo, SessionInfo):
                                    print('Session User ID : ' + str(sessionInfo.get_session_url()))

                                sessionUserInfo = sessionInfoObj.get_user_info()

                                if isinstance(sessionUserInfo, SessionUserInfo):
                                    print('Session User ID : ' + str(sessionUserInfo.get_user_id()))
                                    print('Session Display Name : ' + str(sessionUserInfo.get_display_name()))
                            elif isinstance(sessionInfoObj, InvalidConfigurationException):
                                print('Invalid configuration exception.')
                                print('Error Code  : ' + str(sessionInfoObj.get_code()))
                                print("Error Message : " + str(sessionInfoObj.get_message()))
                                if sessionInfoObj.get_parameter_name() is not None:
                                    print("Error Parameter Name : " + str(sessionInfoObj.get_parameter_name()))
                                if sessionInfoObj.get_key_name() is not None:
                                    print("Error Key Name : " + str(sessionInfoObj.get_key_name()))
                            else:
                                print('Get Session Details Request Failed')
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


GetSessionDetails.execute()
