# This is a sample Python script.
import os

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from zohosdk.src.com.zoho.exception.sdk_exception import SDKException
from zohosdk.src.com.zoho.user_signature import UserSignature
from zohosdk.src.com.zoho.dc.data_center import DataCenter
from zohosdk.src.com.zoho.api.authenticator.api_key import APIKey
from zohosdk.src.com.zoho.util.constants import Constants
from zohosdk.src.com.zoho.api.logger import Logger
from zohosdk.src.com.zoho import Initializer

from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentInfo, UserInfo, Margin, DocumentDefaults, EditorSettings, \
    CallbackSettings, MailMergeTemplateParameters, InvalidConfigurationException
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

import time

from zohosdk.src.com.zoho.util import StreamWrapper


class CreateMergeTemplate:

    @staticmethod
    def execute():
        CreateMergeTemplate.init_sdk()
        createTemplateParams = MailMergeTemplateParameters()

        createTemplateParams.set_url("https://demo.office-integrator.com/zdocs/Graphic-Design-Proposal.docx")
        createTemplateParams.set_merge_data_json_url("https://demo.office-integrator.com/data/candidates.json")

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/OfferLetter.zdoc"
        # print('Source document file path : ' + filePath)
        # createTemplateParams.set_document(StreamWrapper(file_path=filePath))
        #
        # jsonFilePath = ROOT_DIR + "/sample_documents/candidates.json"
        # print('Data Source Json file to be path : ' + filePath)
        # createTemplateParams.set_merge_data_json_content(StreamWrapper(file_path=jsonFilePath))

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("New Document")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        createTemplateParams.set_document_info(documentInfo)

        # Optional Configuration - Add User meta in request to identify the user in document session
        userInfo = UserInfo()
        userInfo.set_user_id("1000")
        userInfo.set_display_name("User 1")

        createTemplateParams.set_user_info(userInfo)

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

        createTemplateParams.set_callback_settings(callbackSettings)

        # Optional Configuration - Set margin while creating document itself.
        # It's applicable only for new documents.
        margin = Margin()

        margin.set_top("1in")
        margin.set_bottom("1in")
        margin.set_left("1in")
        margin.set_right("1in")

        # Optional Configuration - Set default settings for document while creating document itself.
        # It's applicable only for new documents.
        documentDefaults = DocumentDefaults()

        documentDefaults.set_font_size(12)
        documentDefaults.set_paper_size("A4")
        documentDefaults.set_font_name("Arial")
        documentDefaults.set_track_changes("enabled")
        documentDefaults.set_orientation("landscape")

        documentDefaults.set_margin(margin)
        documentDefaults.set_language("ta")

        createTemplateParams.set_document_defaults(documentDefaults)

        # Optional Configuration
        editorSettings = EditorSettings()

        editorSettings.set_unit("in")
        editorSettings.set_language("en")
        editorSettings.set_view("pageview")

        createTemplateParams.set_editor_settings(editorSettings)

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

        createTemplateParams.set_permissions(permissions)

        v1Operations = V1Operations()
        response = v1Operations.create_mail_merge_template(createTemplateParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CreateDocumentResponse):
                    print('Document Id : ' + str(responseObject.get_document_id()))
                    print('Document Session ID : ' + str(responseObject.get_session_id()))
                    print('Document Session URL : ' + str(responseObject.get_document_url()))
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
                    print('Merge Template Creation Request Failed')

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


CreateMergeTemplate.execute()
