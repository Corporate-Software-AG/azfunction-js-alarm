import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';

const TEAMS_WEBHOOK = process.env["TeamsWebHook"]

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    context.log("----------------")
    context.log(req.body)
    context.log("----------------")
    context.log(req.body.value[0].resourceData)
    context.log("------User------")
    context.log(req.body.value[0].resourceData.source.identity.user)
    context.log("---cR-original--")
    context.log(req.body.value[0].resourceData.callRoutes[0].original)
    context.log("---cR-final-----")
    context.log(req.body.value[0].resourceData.callRoutes[0].final)

    await sendToTeams();
    context.log("-----finish-----")

};
export default httpTrigger;

async function sendToTeams() {
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
                        "body": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "width": "stretch",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "ALARM",
                                                "wrap": true,
                                                "size": "Medium",
                                                "weight": "Bolder"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Column",
                                        "width": "stretch",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "BLABLA",
                                                "wrap": true,
                                                "weight": "Lighter"
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "TextBlock",
                                "text": "ALARM SEND!!",
                                "wrap": true
                            }
                        ],
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.3"
                    }
                }
            ]
        }
    }
    await axios(config);
}

