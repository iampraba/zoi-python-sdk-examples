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
            user = UserSignature("john@zylker.com")
            environment = DataCenter.Environment("https://api.office-integrator.com", None, None, None)
            apikey = APIKey("2ae438cf864488657cc9754a27daa480", Constants.PARAMS)
            logger = Logger.get_instance(Logger.Levels.INFO, "./logs.txt")

            Initializer.initialize(user, environment, apikey, None, None, logger, None)

        except SDKException as ex:
            print(ex.code)
