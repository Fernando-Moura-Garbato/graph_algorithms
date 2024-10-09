import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import pkg from 'pg';
const { Client: PgClient } = pkg;
import * as fs from 'node:fs/promises'

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



//**************//
//---REQUESTS---//
//**************//

// Obtendo os nomes de todos os arquivos de um drive
// await graphClient.api('/users/suporte01@grupounus.com.br/drive/root/children')
//             .select('name')
//             .get()
//             .then( (resposta) => {
//                 resposta.value.forEach(item => {
//                     console.log(item.name);
//                 })
//             })

//Obtendo um valor específico
//await graphClient.api('/users/suporte02@grupounus.com.br/messages').get().then( (resposta) => {console.log(resposta.value[5].subject)})

async function folderSearch(user, folderId, counter, search){
    let folder = await graphClient.api('users/' + `${user}` + '/drive/items/' + `${folderId}` + '/children').get()
    for(let i = 0; i < folder.value.length; i++){
        if(folder.value[i].name.includes(search)){
            counter++;
        }
        if('folder' in folder.value[i]){
            folderSearch(user, folder.value[i].id, counter, search)
        }
    }
}

console.log(await graphClient.api('users/suporte02@grupounus.com.br/drive/root/children').get())


async function officeSearch(clientVal) {
    for(let i = 0; i < clientVal.value.length; i++){
        let usuario = clientVal.value[i].userPrincipalName;
        await handle.writeFile(clientVal.value[i].displayName + ';');
        
        await graphClient.api('users/suporte02@grupounus.com.br/drive/root/children').get()

    }
}


let handle = await fs.open('C:/Users/fernando.garbato/Desktop/graph_demo/teste.txt', 'w')
//Paginação de dados
// try{
//     let resposta = await graphClient.api('/users').get();
//     await officeSearch(resposta);
//     let nextPage = resposta['@odata.nextLink'];
//     while (nextPage!=undefined){
//         let respostaProx = await graphClient.api(nextPage).get();
//         await officeSearch(respostaProx);
//         nextPage = respostaProx["@odata.nextLink"]
//     }
// } catch(error){
//     console.log(error);
// }

//Input à database
// let pgInst = new PgClient({
//     user: process.env.PG_USER,
//     host: process.env.PG_HOST,
//     database: process.env.PG_DATABASE,
//     password: process.env.PG_PASS,
//     port: process.env.PG_PORT,
// })

// pgInst.connect().then( (resultado) => {console.log(resultado)})

// await pgInst.query("SELECT * FROM teste1").then( (resultado) => {console.log(resultado)})



//pgInst.end()
console.log('\nFinal.')