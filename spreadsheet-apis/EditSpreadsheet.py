# This is a sample Python script.
import os

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from zohosdk.src.com.zoho.exception.sdk_exception import SDKException
from zohosdk.src.com.zoho.user_signature import UserSignature
from zohosdk.src.com.zoho.dc.data_center import DataCenter
from zohosdk.src.com.zoho.api.authenticator.api_key import APIKey
from zohosdk.src.com.zoho.util import StreamWrapper
from zohosdk.src.com.zoho.util.constants import Constants
from zohosdk.src.com.zoho.api.logger import Logger
from zohosdk.src.com.zoho import Initializer

from zohosdk.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, CreateSheetParameters, DocumentInfo, \
    SheetUserSettings, SheetCallbackSettings, SheetUiOptions, SheetEditorSettings, CreateSheetResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

import time


class EditSpreadsheet:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-sheet-edit-spreadsheet-v1.html
    @staticmethod
    def execute():
        EditSpreadsheet.init_sdk()
        editSheetParams = CreateSheetParameters()

        editSheetParams.set_url('https://demo.office-integrator.com/samples/sheet/Contact_List.xlsx')

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/Contact_List.xlsx"
        # print('Path for source file to be edited : ' + filePath)
        # editSheetParams.set_document(StreamWrapper(file_path=filePath))

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("New Spreadsheet")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        editSheetParams.set_document_info(documentInfo)

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = SheetUserSettings()

        userInfo.set_display_name("User 1")

        editSheetParams.set_user_info(userInfo)

        # Optional Configuration - Add callback settings to configure.
        # how file needs to be received while saving the document
        callbackSettings = SheetCallbackSettings()

        # Optional Configuration - configure additional parameters
        # which can be received along with document while save callback
        saveUrlParams = {}

        saveUrlParams['id'] = '123131'
        saveUrlParams['auth_token'] = '1234'
        # Following $<> values will be replaced by actual value in callback request
        # To know more - https://www.zoho.com/officeintegrator/api/v1/zoho-sheet-create-spreadsheet.html#saveurl_params
        saveUrlParams['extension'] = '$format'
        saveUrlParams['document_name'] = '$filename'

        callbackSettings.set_save_url_params(saveUrlParams)

        # Optional Configuration - configure additional headers
        # which could be received in callback request headers while saving document
        saveUrlHeaders = {}

        saveUrlHeaders['access_token'] = '12dweds32r42wwds34'
        saveUrlHeaders['client_id'] = '12313111'

        callbackSettings.set_save_url_headers(saveUrlHeaders)

        callbackSettings.set_save_format("xlsx")
        callbackSettings.set_save_url(
            "https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157123434d4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286")

        editSheetParams.set_callback_settings(callbackSettings)

        # Optional Configuration
        sheetEditorSettings = SheetEditorSettings()

        sheetEditorSettings.set_language('en')
        sheetEditorSettings.set_country('US')

        editSheetParams.set_editor_settings(sheetEditorSettings)

        # Optional Configuration
        sheetUiOptions = SheetUiOptions()

        sheetUiOptions.set_save_button("show")

        editSheetParams.set_ui_options(sheetUiOptions)

        # Optional Configuration - Configure permission values for session
        # based of you application requirement
        permissions = {}

        permissions["document.export"] = True
        permissions["document.print"] = True
        permissions["document.edit"] = True

        editSheetParams.set_permissions(permissions)

        v1Operations = V1Operations()
        response = v1Operations.create_sheet(editSheetParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateSheetResponse):
                    print('Spreadsheet Id : ' + str(responseObject.get_document_id()))
                    print('Spreadsheet Session ID : ' + str(responseObject.get_session_id()))
                    print('Spreadsheet Session URL : ' + str(responseObject.get_document_url()))
                    print('Spreadsheet Session Grid View URL : ' + str(responseObject.get_gridview_url()))
                    print('Spreadsheet Session Delete URL : ' + str(responseObject.get_session_delete_url()))
                    print('Spreadsheet Delete URL : ' + str(responseObject.get_document_delete_url()))
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Spreadsheet Creation Request Failed')

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


EditSpreadsheet.execute()
