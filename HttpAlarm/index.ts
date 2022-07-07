import { AzureFunction, Context, HttpRequest } from "@azure/functions"

const IOT_DEVICE_CONNECTION_STRING = process.env.IOT_DEVICE_CONNECTION_STRING
const IOT_DEVICE_NAME = process.env.IOT_DEVICE_NAME

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    const appInsights = require('applicationinsights')
    appInsights.setup();
    const appInsightsClient = appInsights.defaultClient;

    const methodParams = {
        methodName: 'onAlarm',
        payload: "!!ALARM!!",
        responseTimeoutInSeconds: 15 // set response timeout as 15 seconds
    };

    const IotClient = require("azure-iothub").Client;
    let iotClient = IotClient.fromConnectionString(IOT_DEVICE_CONNECTION_STRING);

    iotClient.invokeDeviceMethod(IOT_DEVICE_NAME, methodParams, (err, result) => {
        if (err) {
            console.error('Failed to invoke method \'' + methodParams.methodName + '\': ' + err.message);
        } else {
            console.log(methodParams.methodName + ' on ' + IOT_DEVICE_NAME + ':');
            console.log(JSON.stringify(result, null, 2));
            appInsightsClient.trackEvent({ name: "ALARM to " + IOT_DEVICE_NAME })
        }
    });

    context.log("-----finish-----")
    return null;

};
export default httpTrigger;