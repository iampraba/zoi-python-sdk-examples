import time

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, CreateSheetParameters, \
    DocumentInfo, \
    SheetUserSettings, SheetCallbackSettings, SheetUiOptions, SheetEditorSettings, CreateSheetResponse, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations


class CreateSpreadsheet:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-sheet-create-spreadsheet.html
    @staticmethod
    def execute():
        CreateSpreadsheet.init_sdk()
        createSheetParams = CreateSheetParameters()

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("New Spreadsheet")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        createSheetParams.set_document_info(documentInfo)

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = SheetUserSettings()

        userInfo.set_display_name("User 1")

        createSheetParams.set_user_info(userInfo)

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

        createSheetParams.set_callback_settings(callbackSettings)

        # Optional Configuration
        sheetEditorSettings = SheetEditorSettings()

        sheetEditorSettings.set_language('en')
        sheetEditorSettings.set_country('US')

        createSheetParams.set_editor_settings(sheetEditorSettings)

        # Optional Configuration
        sheetUiOptions = SheetUiOptions()

        sheetUiOptions.set_save_button("show")

        createSheetParams.set_ui_options(sheetUiOptions)

        # Optional Configuration - Configure permission values for session
        # based of you application requirement
        permissions = {}

        permissions["document.export"] = True
        permissions["document.print"] = True
        permissions["document.edit"] = True

        createSheetParams.set_permissions(permissions)

        v1Operations = V1Operations()
        response = v1Operations.create_sheet(createSheetParams)

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

CreateSpreadsheet.execute()