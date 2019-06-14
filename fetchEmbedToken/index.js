const axios = require('axios');
const tenantId = process.env.PBIE_TENANT_ID;
const clientSecret = process.env.PBIE_CLIENT_SECRET;
const clientId = process.env.PBIE_CLIENT_ID;
const groupId = process.env.PBIE_GROUP_ID;
const reportId = process.env.PBIE_REPORT_ID;
const username = process.env.PBIE_USERNAME;
const password = process.env.PBIE_PASSWORD;
const powerBiUrl = "https://api.powerbi.com/v1.0/myorg/groups/" + groupId + "/dashboards/" + reportId;
const post = "post";

module.exports = async function(context) {

    async function httpRequest(method, url, headers, data) {
        try {
            const httpResponse = await axios({ method, url, headers, data });
            return httpResponse.data;
        } catch (err) {
            console.log("Http Request Error", err);
        }
    }

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
            const authResponse = await httpRequest(post, authUrl, headers, data);
            return authResponse.access_token;
        } catch (err) {
            context.log("fetchAccessTokenError: ", err);
        }
    }

    async function fetchEmbedUrl(accessToken) {
        const embedUrlHeaders = {
            "Authorization": "Bearer " + accessToken
        };

        try {
            const embedUrlResponse = await httpRequest("get", powerBiUrl, embedUrlHeaders);
            contet.log("embed response", embedUrlResponse);
            return embedUrlResponse.embedUrl;
        } catch (err) {
            context.log("Fetch Embed Url Error: ", err);
        }
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

        try {
            const response = await httpRequest(post, url, embedTokenHeaders, embedTokenBody);
            return response.token;
        } catch (err) {
            context.log("Fetch Embed Token Error: ", err);
        }
    }

    const data = await fetchAccessToken()
        .then(accessToken => {
            const embedUrl = fetchEmbedUrl(accessToken);
            const embedToken = fetchEmbedToken(accessToken);

            return Promise.all([embedUrl, embedToken, accessToken]);
        })
        .then((resp) => {
            context.log("resp", resp);
            const result = { accessToken: resp[2], embedUrl: resp[0], embedToken: resp[1], reportId };
            context.log('Result: ', result);
            return result;
        })
        .catch(err => context.log("Error on main function return: ", err));

    context.res = {
        body: JSON.stringify(data)
    };

    context.done();
};