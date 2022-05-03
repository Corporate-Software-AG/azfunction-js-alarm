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

    let rows = []
    if (user) {
        rows.push(getTableRow("User ID: ", user.id))
        rows.push(getTableRow("Tenant ID: ", user.tenantId))
    }
    if (phone) {
        rows.push(getTableRow("Number: ", phone))
    }

    context.log(rows)
    await sendToTeams(rows);
    context.log("-----finish-----")

};
export default httpTrigger;

async function sendToTeams(rows: object[]) {
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
                                "type": "TextBlock",
                                "size": "Medium",
                                "weight": "Bolder",
                                "text": "ALARM",
                                "horizontalAlignment": "Center"
                            },
                            {
                                "type": "Table",
                                "columns": [
                                    {
                                        "width": 1
                                    },
                                    {
                                        "width": 3
                                    }
                                ],
                                "rows": rows
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

function getTableRow(key: string, value: string): object {
    return {
        "type": "TableRow",
        "cells": [
            {
                "type": "TableCell",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": key,
                        "wrap": true
                    }
                ]
            },
            {
                "type": "TableCell",
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