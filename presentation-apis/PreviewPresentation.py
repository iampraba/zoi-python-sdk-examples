import time
import os

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.util import StreamWrapper
from officeintegrator.src.com.zoho.officeintegrator.v1 import DocumentInfo, InvalidConfigurationException, \
    PreviewResponse, \
    PresentationPreviewParameters, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class PreviewPresentation:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/zoho-show-preview-presentation.html
    @staticmethod
    def execute():
        PreviewPresentation.init_sdk()
        previewParameter = PresentationPreviewParameters()

        # Either use url as document source or attach the document in request body use below methods
        previewParameter.set_url('https://demo.office-integrator.com/samples/show/Zoho_Show.pptx')

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/Zoho_Show.pptx"
        # print('Path for source file to be edited : ' + filePath)
        # previewParameter.set_document(StreamWrapper(file_path=filePath))

        # Optional Configuration - Add document meta in request to identify the file in Zoho Server
        documentInfo = DocumentInfo()
        documentInfo.set_document_name("New Document")
        documentInfo.set_document_id((round(time.time() * 1000)).__str__())

        previewParameter.set_document_info(documentInfo)

        previewParameter.set_language('en')

        v1Operations = V1Operations()
        response = v1Operations.create_presentation_preview(previewParameter)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, PreviewResponse):
                    print('\nPresentation Id : ' + str(responseObject.get_document_id()))
                    print('Presentation Session ID : ' + str(responseObject.get_session_id()))
                    print('Presentation Preview URL : ' + str(responseObject.get_preview_url()))
                    print('Presentation Session Delete URL : ' + str(responseObject.get_session_delete_url()))
                    print('Presentation Delete URL : ' + str(responseObject.get_document_delete_url()))
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Preview Presentation Request Failed')

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

PreviewPresentation.execute()