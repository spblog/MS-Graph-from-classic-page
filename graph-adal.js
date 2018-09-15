(function () {
    let clientId = _spPageContextInfo.spfx3rdPartyServicePrincipalId;

    let authContext = new AuthenticationContext({
        clientId: clientId,
        tenant: _spPageContextInfo.aadTenantId, 
        redirectUri: window.location.origin + '/_forms/spfxsinglesignon.aspx' 
    });

    function silentLogin() {
        return new Promise(function (resolve, reject) {
            authContext._renewToken(clientId, function (message, token) {
                if (!token) {
                    let err = new Error(message);
                    console.log(err);
                    reject(err);
                }
                else {
                    //console.log(token);
                    let user = authContext.getCachedUser();
                    resolve(user);

                }
            }, authContext.RESPONSE_TYPE.ID_TOKEN_TOKEN);
        });
    }

    function getToken(resource) {
        return new Promise(function (resolve, reject) {
            authContext.acquireToken(resource, function (message, token) {
                if (!token) {
                    let err = new Error(message);
                    console.log(err);
                    reject(err);
                }
                else {
                    //console.log(token);
                    resolve(token);
                }
            });
        });
    }

    silentLogin()
        .then(function (user) {
            console.log(user);

            return getToken('https://graph.microsoft.com');
        })
        .then(function (graphAccessToken) {
            return fetch('https://graph.microsoft.com/v1.0/groups', {
                headers: {
                    'Authorization': 'Bearer ' + graphAccessToken
                }
            })
        })
        .then(function (data) {
            return data.json();
        })
        .then(function (result) {
            console.log(result);
        })
        .catch(console.log);
})();