const axios = require('axios');

const tenantId = process.env.PBIE_TENANT_ID;
const clientSecret = process.env.PBIE_CLIENT_SECRET;
const clientId = process.env.PBIE_CLIENT_ID;
const groupId = process.env.PBIE_GROUP_ID;
const reportId = process.env.PBIE_REPORT_ID;
const username = process.env.PBIE_USERNAME;
const password = process.env.PBIE_PASSWORD;
const powerBiUrl = "https://api.powerbi.com/v1.0/myorg/groups/" + groupId + "/dashboards/" + reportId;

async function httpRequest(method, url, headers, data) {
    try {
        const httpResponse = await axios({ method, url, headers, data });
        return httpResponse.data;
    } catch (err) {
        console.log("Http Request Error", err);
    }
}

module.exports = async function(context, req) {

    context.log("req: ", req);
    async function fetchAccessToken() {

        const authUrl = "https://login.microsoftonline.com/" + tenantId + "/oauth2/token";
        const resource = "https://analysis.windows.net/powerbi/api";
        const dataBody = {
            "client_id": clientId,
            "client_secret": clientSecret,
            "grant_type": "password",
            password,
            resource,
            username
        };
        const data = Object.keys(dataBody).map(key => key + '=' + dataBody[key]).join('&');

        const headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json, text/plain, */*'
        };

        try {
            const authResponse = await httpRequest("post", authUrl, headers, data);
            tokenAquired = true;
            return authResponse.access_token;
        } catch (err) {
            console.log("fetchAccessTokenError: ", err);
        }
    }

    async function fetchEmbedUrl(accessToken) {
        const embedUrlHeaders = {
            "Authorization": "Bearer " + accessToken
        };

        const response = await httpRequest("get", powerBiUrl, embedUrlHeaders);
        // console.log("response -->", response)
        return response.embedUrl;

    }

    async function fetchEmbedToken(accessToken) {
        const embedTokenHeaders = {
            "Authorization": "Bearer " + accessToken,
            "Content-Type": "application/json"
        };
        const url = powerBiUrl + "/GenerateToken";
        const embedTokenBody = {
            accessLevel: "view"
        };
        const response = await httpRequest("post", url, embedTokenHeaders, embedTokenBody);
        // console.log("embed token", response);
        return response.token;

    }

    return await fetchAccessToken()
        .then(token => {
            const embedUrl = fetchEmbedUrl(token);
            // console.log("embedUrl: ", embedUrl);
            const embedToken = fetchEmbedToken(token);
            // console.log("embedToken: ", embedToken);
            return Promise.all([embedUrl, embedToken]);
        })
        .then((resp) => {
            return { embedUrl: resp[0], embedToken: resp[1], reportId };
        })
        .catch(err => console.log("Error on main function return: ", err));

};