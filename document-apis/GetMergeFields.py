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

from zohosdk.src.com.zoho.officeintegrator.v1 import DocumentMeta, InvalidConfigurationException, \
    GetMergeFieldsParameters, MergeFieldsResponse, MergeFields
from zohosdk.src.com.zoho.officeintegrator.v1.create_document_response import CreateDocumentResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations


class GetMergeFields:

    @staticmethod
    def execute():
        GetMergeFields.init_sdk()
        getMergeFilesParams = GetMergeFieldsParameters()

        getMergeFilesParams.set_file_url("https://demo.office-integrator.com/zdocs/OfferLetter.zdoc");

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


GetMergeFields.execute()
