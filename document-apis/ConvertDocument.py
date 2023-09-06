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
from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentConversionParameters, DocumentConversionOutputOptions, \
    FileBodyWrapper, InvalidConfigurationException
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations
from zohosdk.src.com.zoho.util import StreamWrapper


class ConvertDocument:

    @staticmethod
    def execute():
        ConvertDocument.init_sdk()
        documentConversionParameters = DocumentConversionParameters()

        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        filePath = ROOT_DIR + "/sample_documents/Graphic-Design-Proposal.docx"
        print('Source file to be converted : ' + filePath)
        documentConversionParameters.set_document(StreamWrapper(file_path=filePath))

        # documentConversionParameters.set_url('https://demo.office-integrator.com/zdocs/LabReport.zdoc')

        outputOptions = DocumentConversionOutputOptions()

        outputOptions.set_format("pdf")
        outputOptions.set_document_name("conversion_output.pdf")
        outputOptions.set_include_comments("all")
        outputOptions.set_include_changes("all")

        documentConversionParameters.set_output_options(outputOptions)
        documentConversionParameters.set_password("***")

        v1Operations = V1Operations()
        response = v1Operations.convert_document(documentConversionParameters)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, FileBodyWrapper):
                    convertedDocument = responseObject.get_file()

                    if isinstance(convertedDocument, StreamWrapper):
                        outputFileStream = convertedDocument.get_stream()
                        outputFilePath = ROOT_DIR + "/sample_documents/conversion_output.pdf"

                        with open(outputFilePath, 'wb') as outputFileObj:
                            # while True:
                            #     # Read a chunk of data from the input stream
                            #     chunk = outputFileStream.read(1024)  # You can adjust the chunk size as needed
                            #
                            #     # If no more data is read, break the loop
                            #     if not chunk:
                            #         break
                            #
                            #     # Write the chunk of data to the file
                            #     outputFileObj.write(chunk)
                            outputFileObj.write(outputFileStream.content)

                        print("\nCheck converted output file in file path - " + outputFilePath)
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Document Conversion Request Failed')

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


ConvertDocument.execute()
