import { AzureFunction, Context } from "@azure/functions"

const IOT_DEVICE_CONNECTION_STRING = process.env.IOT_DEVICE_CONNECTION_STRING
const IOT_DEVICE_NAME = process.env.IOT_DEVICE_NAME
const IS_CONNECTED = process.env.IS_CONNECTED

const timerTrigger: AzureFunction = async function (context: Context, myTimer: any): Promise<void> {
    var timeStamp = new Date().toISOString();
    if (myTimer.isPastDue) {
        context.log('Timer function is running late!');
    }

    const appInsights = require('applicationinsights')
    appInsights.setup();
    const appInsightsClient = appInsights.defaultClient;

    const methodParams = {
        methodName: 'onHealthCheck',
        payload: "healthcheck",
        responseTimeoutInSeconds: 15 // set response timeout as 15 seconds
    };

    const IotClient = require("azure-iothub").Client;
    let iotClient = IotClient.fromConnectionString(IOT_DEVICE_CONNECTION_STRING);

    iotClient.invokeDeviceMethod(IOT_DEVICE_NAME, methodParams, (err, result) => {
        if (err) {
            console.error('Failed to invoke method \'' + methodParams.methodName + '\': ' + err.message);
            if (IS_CONNECTED === "true") {
                appInsightsClient.trackEvent({ name: "CONNECTION LOST to " + IOT_DEVICE_NAME })
            } else {
                appInsightsClient.trackEvent({ name: "Disabled: " + IOT_DEVICE_NAME })
            }
        } else {
            console.log(methodParams.methodName + ' on ' + IOT_DEVICE_NAME + ':');
            console.log(JSON.stringify(result, null, 2));
            appInsightsClient.trackEvent({ name: 'successful connection to ' + IOT_DEVICE_NAME })
        }
    });

    context.log('Timer trigger function ran!', timeStamp);
};

export default timerTrigger;
