import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";

//Configuração do MSAL
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: process.env.AUTHORITY,
        clientSecret: process.env.CLIENT_SECRET
    }
};

//ensureScope() - Função para admitir novas permissões, desnecessária ao momento
/*
function ensureScope (scope) {
    if (!msalRequest.scopes.some((s) => s.toLowerCase() === scope.toLowerCase())) {
        msalRequest.scopes.push(scope);
    }
}
*/

//Inicializa um cliente daemon do MSAL usando msalConfig
const msalClient = new ConfidentialClientApplication(msalConfig);

//Com o cliente ativo, cria uma request pro Graph e adquire um token
const daemonRequest = {
    scopes: ["https://graph.microsoft.com/.default"]
}
const tentToken = await msalClient.acquireTokenByClientCredential(daemonRequest)

//Adquire o token
async function getToken(){
    return tentToken;
}

// Middleware
const authProvider = {
    getAccessToken: async () => {
        const tentToken = await getToken();
        return tentToken.accessToken;
    }
};

// Inicializa o cliente do graph com Middleware
const graphClient = Client.initWithMiddleware({authProvider});
//É válido relembrar que as respostas do servidor geralmente são objetos preenchindos por diversos metadados.
//Para validar objetos e variáveis, é sempre bom logar.

//Requests
await graphClient.api('/users')
            .select('id,displayName,mail')
            .top(5)
            .get()
            .then( (resposta) => {
                console.log(resposta['@odata.nextLink'])
            })



async function getUsers(){
    for(let i = 0; i++;){
          await graphClient.api('/users')
            .select('id,displayName,mail')
            .top(5)
            .get()
            .then( (resposta) => {
                console.log(resposta['@odata.nextLink'])
            })
    }
}


/*
    //REQUESTS
    async function listUsers() {
        try {
            const users = await graphClient
                .api('/users') // Graph API endpoint to list all users
                .select('id,displayName,mail') // Select fields you want to retrieve
                .get();
            return users.value; // Return the array of users
        } catch (error) {
            console.error("Error fetching users:", error);
        }
    }

    let usuarios = listUsers()
    console.log(usuarios)
*/




console.log('\nFinal.')