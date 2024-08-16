from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer

from officeintegrator.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, PlanDetails
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations


class GetPlanDetails:

    # Refer API documentation - https://www.zoho.com/officeintegrator/api/v1/get-plan-details.html
    @staticmethod
    def execute():
        GetPlanDetails.init_sdk()

        v1Operations = V1Operations()
        response = v1Operations.get_plan_details()

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, PlanDetails):
                    print("\nPlan name - " + str(responseObject.get_plan_name()))
                    print("API usage limit - " + str(responseObject.get_usage_limit()))
                    print("API usage so far - " + str(responseObject.get_total_usage()))
                    print("Plan upgrade payment link - " + str(responseObject.get_payment_link()))
                    print("Subscription period - " + str(responseObject.get_subscription_period()))
                    print("Subscription interval - " + str(responseObject.get_subscription_interval()))
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Get Plan Details Request Failed')

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


GetPlanDetails.execute()
