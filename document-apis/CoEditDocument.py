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
    EditorSettings, UiOptions
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_parameters import CreateDocumentParameters
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

import time
import os

from zohosdk.src.com.zoho.util import StreamWrapper


class CoEditDocument:

    @staticmethod
    def execute():
        CoEditDocument.init_sdk()
        createDocumentParams = CreateDocumentParameters()

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("New Document")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        createDocumentParams.set_document_info(documentInfo)

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1000")
        userInfo.set_display_name("User 1")

        createDocumentParams.set_user_info(userInfo)

        # Optional Configuration - Add callback settings to configure.
        # how file needs to be received while saving the document
        callbackSettings = CallbackSettings()

        # Optional Configuration - configure additional parameters
        # which can be received along with document while save callback
        saveUrlParams = {}

        saveUrlParams['id'] = '123131'
        saveUrlParams['auth_token'] = '1234'
        # Following $<> values will be replaced by actual value in callback request
        # To know more - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-create-document.html#saveurl_params
        saveUrlParams['extension'] = '$format'
        saveUrlParams['document_name'] = '$filename'
        saveUrlParams['session_id'] = '$session_id'

        callbackSettings.set_save_url_params(saveUrlParams)

        # Optional Configuration - configure additional headers
        # which could be received in callback request headers while saving document
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

        # Optional Configuration - Set default settings for document while creating document itself.
        # It's applicable only for new documents.
        documentDefaults = DocumentDefaults()

        documentDefaults.set_track_changes("enabled")
        documentDefaults.set_language("ta")

        createDocumentParams.set_document_defaults(documentDefaults)

        # Optional Configuration
        editorSettings = EditorSettings()

        editorSettings.set_unit("in")
        editorSettings.set_language("en")
        editorSettings.set_view("pageview")

        createDocumentParams.set_editor_settings(editorSettings)

        # Optional Configuration
        uiOptions = UiOptions()

        uiOptions.set_dark_mode("show")
        uiOptions.set_file_menu("show")
        uiOptions.set_save_button("show")
        uiOptions.set_chat_panel("show")

        createDocumentParams.set_ui_options(uiOptions)

        # Optional Configuration - Configure permission values for session
        # based of you application requirement
        permissions = {}

        permissions["document.export"] = True
        permissions["document.print"] = True
        permissions["document.edit"] = True
        permissions["review.comment"] = True
        permissions["review.changes.resolve"] = True
        permissions["document.pausecollaboration"] = True
        permissions["document.fill"] = False

        createDocumentParams.set_permissions(permissions)

        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        filePath = ROOT_DIR + "/sample_documents/Graphic-Design-Proposal.docx"

        print('Path for source file to be edited : ' + filePath)

        createDocumentParams.set_document(StreamWrapper(file_path=filePath))

        # createDocumentParams.set_url('https://demo.office-integrator.com/zdocs/LabReport.zdoc')

        # Creating session1 for collaboration demo
        v1Operations = V1Operations()
        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            response_object = response.get_object()

            if response_object is not None:
                if isinstance(response_object, CreateDocumentResponse):
                    print('Document Id : ' + str(response_object.get_document_id()))
                    print('Document Session 1 ID : ' + str(response_object.get_session_id()))
                    print('Document Session 1 URL : ' + str(response_object.get_document_url()))
                    print('Document Session 1 Delete URL : ' + str(response_object.get_session_delete_url()))
                    print('Document Delete URL : ' + str(response_object.get_document_delete_url()))

        # Creating session2 for collaboration demo
        # Need to use same Document meta to create session for same document again. So modify only other parameters
        # For demo purpose only same configuration object used. Also, two document session created in same request.

        # Add User meta to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1001")
        userInfo.set_display_name("User 2")

        createDocumentParams.set_user_info(userInfo)

        response = v1Operations.create_document(createDocumentParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            response_object = response.get_object()

            if response_object is not None:
                if isinstance(response_object, CreateDocumentResponse):
                    print('Document Id : ' + str(response_object.get_document_id()))
                    print('Document Session 2 ID : ' + str(response_object.get_session_id()))
                    print('Document Session 2 URL : ' + str(response_object.get_document_url()))
                    print('Document Session 2 Delete URL : ' + str(response_object.get_session_delete_url()))
                    print('Document Delete URL : ' + str(response_object.get_document_delete_url()))

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


CoEditDocument.execute()
