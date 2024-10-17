import * as fs from 'node:fs/promises';
import {graphClient} from './auth.js';

const now = new Date();
const daysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
const reportDate = daysAgo.toISOString().slice(0, 19) + 'Z';

let emailUseHandle = await fs.open('C:/Users/fernando.garbato/Desktop/graph_demo/generated/email-use_' + `${reportDate.slice(0,10)}` + '.csv', 'w');

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
        await emailUseReport(await graphClient.api(call["@odata.nextLink"]).select("userPrincipalName, displayName, mail, surname, userType").get());
    }
}

 let emailUseReportCall = await graphClient.api("users").select("userPrincipalName, displayName, mail, surname, userType").get();
 await emailUseReport(emailUseReportCall);

await emailUseHandle.close();

let emailUseRead = await fs.open('C:/Users/fernando.garbato/Desktop/graph_demo/generated/email-use_' + `${reportDate.slice(0,10)}` + '.csv', 'r');
let buffer = Buffer.alloc((await emailUseRead.stat()).size);
let {bytesRead} = await emailUseRead.read(buffer, 0, buffer.length, 0);

const sendEmail = {
    message:{
        subject: "Relatório de uso de e-mail " + `${reportDate.slice(0,10)}`,
        body: {
            contentType: 'HTML',
            content: '<center><h1>Relatório semanal de utilização de e-mail do Grupo Unus</h1></center><h2>Atenção</h2>\nPara que a tabela seja formatada da maneira correta, abra uma planilha em branco, vá na aba Dados -> \"De Text/CSV\" e importe o arquivo.'
        },
        toRecipients: [
            {
                emailAddress: {
                    address: 'suporte02@grupounus.com.br'
                }
            }
        
        ],
        attachments:[
            {
              '@odata.type': '#microsoft.graph.fileAttachment',
              name: 'relatorio_email_' + `${reportDate.slice(0,10)}` + '.csv',
              contentType: 'text/plain',
              contentBytes: buffer.toString('base64')
            }
          ]
    },
    saveToSentItems: 'true'
}


await graphClient.api("users/automacoes@grupounus.com.br/sendMail").post(sendEmail);

// + A fazer: remoção automática de arquivos antigos.
emailUseRead.close();