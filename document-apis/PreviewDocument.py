import os

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.util import StreamWrapper
from officeintegrator.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, PreviewParameters, \
    PreviewDocumentInfo, PreviewResponse, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class PreviewDocument:

    # Refer Preview API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-writer-preview-document.html
    @staticmethod
    def execute():
        PreviewDocument.init_sdk()
        previewParameter = PreviewParameters()

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = PreviewDocumentInfo()
        documentInfo.set_document_name("New Document")

        previewParameter.set_document_info(documentInfo)

        # Optional Configuration - Configure permission values for session
        # based of you application requirement
        permissions = {}

        permissions["document.print"] = True

        previewParameter.set_permissions(permissions)

        # Either use url as document source or attach the document in request body use below methods
        previewParameter.set_url('https://demo.office-integrator.com/zdocs/LabReport.zdoc')

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/Graphic-Design-Proposal.docx"
        # print('Path for source file to be edited : ' + filePath)
        # previewParameter.set_document(StreamWrapper(file_path=filePath))

        v1Operations = V1Operations()
        response = v1Operations.create_document_preview(previewParameter)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, PreviewResponse):
                    print('Document Id : ' + str(responseObject.get_document_id()))
                    print('Document Session ID : ' + str(responseObject.get_session_id()))
                    print('Document Preview URL : ' + str(responseObject.get_preview_url()))
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
                    print('Preview Document Request Failed')

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

PreviewDocument.execute()