import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';

const APP_ID = process.env["appId"];
const APP_SECRET = process.env["appSecret"];
const TENANT_ID = process.env["tenantId"];
const TEAMS_WEBHOOK = process.env["teamsWebhook"]

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    context.log(req.body)
    //await sendToTeams(req.body);

};
export default httpTrigger;

async function sendToTeams(body: any) {
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
                                                "text": body.name,
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
                                                "text": body.email,
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

