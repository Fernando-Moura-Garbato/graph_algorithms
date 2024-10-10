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


//This function grabs a user, a folder ID and a search term, creates an API call and passes the result through an algorithm to count how many
//times the search term appears on file names. It also recursively searches through folders within that folder, and in both cases, works with
//data pagination. The repeatNext() function is declared within merely for organization.
async function folderSearch(user, folderId, search){
    let counter = 0;
    let folder = await graphClient.api('users/' + `${user}` + '/drive/items/' + `${folderId}` + '/children').get()
    for(let i = 0; i < folder.value.length; i++){
        if(folder.value[i].name.includes(search)){
            counter++;
        }
        if('folder' in folder.value[i]){
            counter = counter + await folderSearch(user, folder.value[i].id, search)
        }
        async function repeatNext(folder, search){
                let counterNext = 0;
                let chamada = await graphClient.api(folder['@odata.nextLink']).get();
                for(let i = 0; i < chamada.value.length; i++){
                    if(chamada.value[i].name.includes(search)){
                        counterNext++;
                    }
                    if('folder' in chamada.value[i]){
                        counterNext = counterNext + await folderSearch(user, chamada.value[i].id, search)
                    }
                    if(i === chamada.value.length - 1 && chamada['@odata.nextLink']){
                        counterNext = counterNext + await repeatNext(chamada['@odata.nextLink'], search)
                    }            
                }
                return counterNext;
        }
        if(i === folder.value.length - 1 && folder['@odata.nextLink']){
        counter = counter + await repeatNext(folder, search)
        }
    }
    return counter;
}
//This function looks within each message body preview in search for invites, and then 
async function emailSearch(call, user){
    let counter = {csv:0, docx:0, xlsx:0};
    for(let i = 0; i < call.value.length; i++){
        if(call.value[i].bodyPreview.includes('invited you to edit a file')){
            let messageWithHtml = await graphClient.api('users/' + `${user}` + '/messages/' + `${call.value[i].id}`).get();
            if(messageWithHtml.body.content.includes('csv')){
                counter.csv++
            }
            if(messageWithHtml.body.content.includes('docx')){
                counter.docx++
            }
            if(messageWithHtml.body.content.includes('xlsx')){
                counter.xlsx++
            }
        }
    }
    if(call['@odata.nextLink']){
        let nextCall = await graphClient.api(call['@odata.nextLink']).get();
        let nextSearch = await emailSearch(nextCall, user);
        counter.csv = counter.csv + emailSearch.csv
        counter.docx = counter.docx + emailSearch.docx
        counter.xlsx = counter.xlsx + emailSearch.xlsx
    }
    return counter;
}

let call1 = await graphClient.api('users/suporte02@grupounus.com.br/messages').get();
console.log(await emailSearch(call1, 'suporte02@grupounus.com.br'));



async function officeSearch(usuarios) {
    for(let i = 0; i < usuarios.value.length; i++){
        let usuario = usuarios.value[i].userPrincipalName;
        await handle.writeFile(usuarios.value[i].displayName + ';');

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