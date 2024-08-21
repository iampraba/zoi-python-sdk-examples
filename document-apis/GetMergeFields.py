import os

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.util import StreamWrapper
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, \
    GetMergeFieldsParameters, MergeFieldsResponse, MergeFields, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class GetMergeFields:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/get-list-of-fields-in-the-document.html
    @staticmethod
    def execute():
        GetMergeFields.init_sdk()
        getMergeFilesParams = GetMergeFieldsParameters()

        # Either use url as document source or attach the document in request body use below methods
        getMergeFilesParams.set_file_url("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc")

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/OfferLetter.zdoc"
        # print('Path for source file to be edited : ' + filePath)
        # getMergeFilesParams.set_file_content(StreamWrapper(file_path=filePath))

        v1Operations = V1Operations()
        response = v1Operations.get_merge_fields(getMergeFilesParams)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, MergeFieldsResponse):
                    mergeFieldsObj = responseObject.get_merge()

                    if isinstance(mergeFieldsObj, list):
                        print('\n---- Total Fields in Document : ' + str(len(mergeFieldsObj)) + " ----")
                        for mergeFieldObj in mergeFieldsObj:
                            if isinstance(mergeFieldObj, MergeFields):
                                print('\nMerge Field ID : ' + mergeFieldObj.get_id())
                                print('Merge Field Display Name : ' + mergeFieldObj.get_display_name())
                                print('Merge Field Type : ' + mergeFieldObj.get_type())

                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Create Merge Fields Request Failed')

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

GetMergeFields.execute()