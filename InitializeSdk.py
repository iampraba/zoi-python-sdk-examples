from zohosdk.src.com.zoho.exception.sdk_exception import SDKException
from zohosdk.src.com.zoho.user_signature import UserSignature
from zohosdk.src.com.zoho.dc.data_center import DataCenter
from zohosdk.src.com.zoho.api.authenticator.api_key import APIKey
from zohosdk.src.com.zoho.util.constants import Constants
from zohosdk.src.com.zoho.api.logger import Logger
from zohosdk.src.com.zoho import Initializer


class InitializeSdk(object):

    @staticmethod
    def execute(self):
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
