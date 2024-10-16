import * as fs from 'node:fs/promises'
import {graphClient} from './auth.js';
import { report } from 'node:process';

const now = new Date();
const daysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
const reportDate = daysAgo.toISOString().slice(0, 19) + 'Z'

let emailUseHandle = await fs.open('C:/Users/fernando.garbato/Desktop/graph_demo/generated/email_use.csv', 'w')

await emailUseHandle.writeFile("Nome;Email;Tipo;Enviados;Recebidos\n");
async function emailUseReport(call){
    for(let i = 0; i < call.value.length; i++){
    let tipo = "";
    if(call.value[i].userType == "Guest"){
        tipo = "Externo"
    } else if(call.value[i].surname){
        tipo = "Usuário"
    } else {
        tipo = "Caixa compartilhada"
    }
    try{
        await emailUseHandle.write(
            `${call.value[i].displayName}` + ';' + 
            `${call.value[i].mail}` + ';' + 
            `${tipo}` + ';'
        );
        let emailUseData = await graphClient.api("users/" + `${call.value[i].userPrincipalName}` + "/messages")
        .select("id")
        .filter("receivedDateTime ge " + `${reportDate}`)
        .count(true)
        .get();
        let emailSentData = await graphClient.api("users/" + `${call.value[i].userPrincipalName}` + "/mailFolders/sentItems/messages")
        .select("id")
        .filter("sentDateTime ge " + `${reportDate}`)
        .count(true)
        .get();
        await emailUseHandle.writeFile(String( `${emailUseData["@odata.count"] - emailSentData["@odata.count"]}`) + ';');
        await emailUseHandle.writeFile(String(emailSentData["@odata.count"]) + "\n");
    } catch(error){
        console.log(error);
        await emailUseHandle.writeFile("\n");
        continue;
    }
    }
    if(call['@odata.nextLink']){
        emailUseReport(await graphClient.api(call["@odata.nextLink"]).select("userPrincipalName, displayName, mail, surname, userType").get());
    }
}

console.log(reportDate.slice(0,10))

const message = {
    subject: "Relatório semanal" + `${reportDate.slice(0,10)}`,
    importance: 'Low',
    body: {
        contentType: 'HTML',
        content: '<h1>Atenção</h1>\nPara que a tabela seja formatada da maneira correta, abra um planilha em branco, vá na aba Dados -> \"De Text/CSV\" e importe o arquivo.'
    },
    toRecipients: [
        {
            emailAddress: {
                address: 'suporte02@grupounus.com.br'
            }
        }
    ]
};




//   let emailUseReportCall = await graphClient.api("users").select("userPrincipalName, displayName, mail, surname, userType").get()
//   emailUseReport(emailUseReportCall);