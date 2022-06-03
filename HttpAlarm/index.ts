import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';

const TEAMS_WEBHOOK = process.env["TeamsWebHook"]
const IOTHUB_CONNECTION_STRING = process.env.IOTHUB_CONNECTION_STRING
const IOT_NAME = "DemoPi"


const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    let user = req.body.value[0].resourceData.source.identity.user
    let phone = req.body.value[0].resourceData.source.identity.phone

    let elements = []

    elements.push({
        "type": "TextBlock",
        "size": "ExtraLarge",
        "weight": "Bolder",
        "text": "!!! ALARM !!!",
        "color": "Attention",
        "horizontalAlignment": "Center"
    });

    if (user) {
        context.log(user)
        elements.push(getRow("User ID", user.id))
        elements.push(getRow("Tenant ID", user.tenantId))
    }
    if (phone) {
        context.log(phone)
        elements.push(getRow("Number", phone.id))
    }

    //await sendToTeams(elements);

    const methodParams = {
        methodName: 'onAlarm',
        payload: "!!ALARM!!",
        responseTimeoutInSeconds: 15 // set response timeout as 15 seconds
    };

    const IotClient = require("azure-iothub").Client;
    let iotClient = IotClient.fromConnectionString(IOTHUB_CONNECTION_STRING);

    iotClient.invokeDeviceMethod(IOT_NAME, methodParams, (err, result) => {
        if (err) {
            console.error('Failed to invoke method \'' + methodParams.methodName + '\': ' + err.message);
        } else {
            console.log(methodParams.methodName + ' on ' + IOT_NAME + ':');
            console.log(JSON.stringify(result, null, 2));
        }
    });

    context.log("-----finish-----")

};
export default httpTrigger;

async function sendToTeams(body: object[]) {
    let config: AxiosRequestConfig = {
        method: 'post',
        url: TEAMS_WEBHOOK,
        headers: {
            'ContentType': 'Application/Json'
        },
        data: {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                        "type": "AdaptiveCard",
                        "body": body,
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.5"
                    }
                }
            ]
        }
    }
    await axios(config);
}

function getRow(key: string, value: string): object {
    return {
        "type": "ColumnSet",
        "columns": [
            {
                "type": "Column",
                "width": 30,
                "items": [
                    {
                        "type": "TextBlock",
                        "text": key,
                        "wrap": true
                    }
                ]
            },
            {
                "type": "Column",
                "width": 70,
                "items": [
                    {
                        "type": "TextBlock",
                        "text": value,
                        "wrap": true
                    }
                ]
            }
        ]
    }
}