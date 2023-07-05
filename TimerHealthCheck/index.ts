import { AzureFunction, Context } from "@azure/functions"
import { Registry, Client } from "azure-iothub";

const IOT_DEVICE_CONNECTION_STRING = process.env.IOTHUB_CONNECTION_STRING
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

    const registry = Registry.fromConnectionString(IOT_DEVICE_CONNECTION_STRING);
    const client = Client.fromConnectionString(IOT_DEVICE_CONNECTION_STRING);
    let devices = (await registry.list()).responseBody;
    for (let d of devices) {
        client.invokeDeviceMethod(d.deviceId, methodParams, (err, result) => {
            if (err) {
                console.error('Failed to invoke method \'' + methodParams.methodName + '\': ' + err.message);
                if (IS_CONNECTED === "true") {
                    appInsightsClient.trackEvent({ name: "CONNECTION LOST to " + d.deviceId })
                } else {
                    appInsightsClient.trackEvent({ name: "Disabled: " + d.deviceId })
                }
            } else {
                console.log(methodParams.methodName + ' on ' + d.deviceId + ':');
                console.log(JSON.stringify(result, null, 2));
                appInsightsClient.trackEvent({ name: 'successful connection to ' + d.deviceId })
            }
        });
    }
    context.log('Timer trigger function ran!', timeStamp);
};

export default timerTrigger;
