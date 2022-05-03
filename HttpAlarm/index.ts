import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';

const TEAMS_WEBHOOK = process.env["TeamsWebHook"]

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    context.log("------Caller------")
    let user = req.body.value[0].resourceData.source.identity.user
    context.log(user)
    let phone = req.body.value[0].resourceData.source.identity.phone
    context.log(phone)

    let elements = []

    rows.push({
        "type": "TextBlock",
        "size": "Medium",
        "weight": "Bolder",
        "text": "ALARM",
        "horizontalAlignment": "Center"
    });

    if (user) {
        rows.push(getRow("User ID", user.id))
        rows.push(getRow("Tenant ID", user.tenantId))
    }
    if (phone) {
        rows.push(getRow("Number: ", phone))
    }

    context.log(rows)
    await sendToTeams(rows);
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
        "type": "TextBlock",
        "text": key + " - " + value,
        "wrap": true
    }
}