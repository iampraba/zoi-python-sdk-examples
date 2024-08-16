import os

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer

from officeintegrator.src.com.zoho.officeintegrator.v1 import CompareDocumentParameters, CompareDocumentResponse, \
    InvalidConfigurationException, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations
from officeintegrator.src.com.zoho.officeintegrator.util import StreamWrapper

class CompareDocument:

    # Refer Compare API documentation - https://www.zoho.com/officeintegrator/api/v1/writer-comparison-api.html
    @staticmethod
    def execute():
        CompareDocument.init_sdk()
        compareDocumentParameters = CompareDocumentParameters();

        # Either use url as document source or attach the document in request body use below methods
        compareDocumentParameters.set_url1('https://demo.office-integrator.com/zdocs/MS_Word_Document_v0.docx')
        compareDocumentParameters.set_url2('https://demo.office-integrator.com/zdocs/MS_Word_Document_v1.docx')

        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

        file1Name = "MS_Word_Document_v0.docx"
        # file1Path = ROOT_DIR + "/sample_documents/" + file1Name
        # print('source file1 path : ' + file1Path)
        # compareDocumentParameters.set_document1(StreamWrapper(file_path=file1Path))

        file2Name = "MS_Word_Document_v1.docx"
        # file2Path = ROOT_DIR + "/sample_documents/" + file2Name
        # print('source file1 path : ' + file2Path)
        # compareDocumentParameters.set_document2(StreamWrapper(file_path=file2Path))

        compareDocumentParameters.set_lang("en")
        compareDocumentParameters.set_title(file1Name + " vs " + file2Name)

        v1Operations = V1Operations()
        response = v1Operations.compare_document(compareDocumentParameters)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, CompareDocumentResponse):
                    print('Document Compare Session URL  : ' + str(responseObject.get_compare_url()))
                    print('Document Compare Session Delete URL : ' + str(responseObject.get_session_delete_url()))
            elif isinstance(responseObject, InvalidConfigurationException):
                print('Invalid configuration exception.')
                print('Error Code  : ' + str(responseObject.get_code()))
                print("Error Message : " + str(responseObject.get_message()))
                if responseObject.get_parameter_name() is not None:
                    print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                if responseObject.get_key_name() is not None:
                    print("Error Key Name : " + str(responseObject.get_key_name()))
            else:
                print('Compare Document Request Failed')

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


CompareDocument.execute()
