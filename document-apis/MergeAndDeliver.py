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

from zohosdk.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, \
    MergeAndDeliverViaWebhookParameters, MailMergeWebhookSettings, MergeAndDeliverViaWebhookSuccessResponse
from zohosdk.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations


class MergeAndDeliver:

    @staticmethod
    def execute():
        MergeAndDeliver.init_sdk()
        parameter = MergeAndDeliverViaWebhookParameters()

        parameter.set_file_url('https://demo.office-integrator.com/zdocs/OfferLetter.zdoc')

        # ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/OfferLetter.zdoc"
        # print('Source document file path : ' + filePath)
        # parameter.set_file_content(StreamWrapper(file_path=filePath))

        # filePath = ROOT_DIR + "/sample_documents/csv_data_source.csv"
        # print('Source document file path : ' + filePath)
        # parameter.set_merge_data_csv_content(StreamWrapper(file_path=filePath))
        # parameter.set_merge_data_csv_url('https://demo.office-integrator.com/data/csv_data_source.csv')

        parameter.set_merge_data_json_url('https://demo.office-integrator.com/data/candidates.json')

        # jsonFilePath = ROOT_DIR + "/sample_documents/candidates.json"
        # print('Data Source Json file to be path : ' + jsonFilePath)
        # parameter.set_merge_data_json_content(StreamWrapper(file_path=jsonFilePath))

        parameter.set_merge_to('separatedoc')
        parameter.set_output_format('pdf')
        parameter.set_password('***')

        webhookSettings = MailMergeWebhookSettings()

        webhookSettings.set_invoke_url('https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286')
        webhookSettings.set_invoke_period('oncomplete')

        parameter.set_webhook(webhookSettings)

        v1Operations = V1Operations()
        response = v1Operations.merge_and_deliver_via_webhook(parameter)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, MergeAndDeliverViaWebhookSuccessResponse):
                    mergeReportUrl = responseObject.get_merge_report_data_url()
                    mergeRecords = responseObject.get_records()

                    print('\nMerge Report URL : ' + mergeReportUrl)

                    if isinstance(mergeRecords, list):
                        print('\n---- Total Merged Records Count : ' + str(len(mergeRecords)) + " ----")
                        for records in mergeRecords:
                            print('Records : ' + str(records))
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


MergeAndDeliver.execute()
