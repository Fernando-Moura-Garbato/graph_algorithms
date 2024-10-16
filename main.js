import pkg from 'pg';
const { Client: PgClient } = pkg;
import * as fs from 'node:fs/promises'

import {graphClient} from './auth.js';

//**************//
//---REQUESTS---//
//**************//

//This function grabs a user, a folder ID and a search term, creates an API call and passes the result through an algorithm to count how many
//times the search term appears on file names. It also recursively searches through folders within that folder, and in both cases, works with
//data pagination. The repeatNext() function is declared within merely for organization.
async function folderSearch(user, folder){
    let counter = {csv:0, docx:0, xlsx:0, csvSize:0, docxSize:0, xlsxSize:0};

    for(let i = 0; i < folder.value.length; i++){
    
        if(folder.value[i].name.includes('csv')){
            counter.csv++;
            counter.csvSize += folder.value[i].size;
        }
        if(folder.value[i].name.includes('docx')){
            counter.docx++;
            counter.docxSize += folder.value[i].size;
        }
        if(folder.value[i].name.includes('xlsx')){
            counter.xlsx++;
            counter.xlsxSize += folder.value[i].size;
        }


        if('folder' in folder.value[i]){
            let folderCall = await graphClient.api('users/' + `${user}` + '/drive/items/' + `${folder.value[i].id}` + '/children').top(100).select('file,folder,name,id,size').get();
            let folderCallResult = await folderSearch(user, folderCall);
            counter.csv += folderCallResult.csv;
            counter.csvSize += folderCallResult.csvSize;
            counter.docx += folderCallResult.docx;
            counter.docxSize += folderCallResult.docxSize;
            counter.xlsx += folderCallResult.xlsx;
            counter.xlsxSize += folderCallResult.xlsxSize;
            }
    }

    if(folder['@odata.nextLink']){
        let nextCall = await graphClient.api(folder['@odata.nextLink']).top(100).select('file,folder,name,id,size').get();
        let nextCallResult = await folderSearch(user, nextCall);
        counter.csv += nextCallResult.csv;
        counter.csvSize += nextCallResult.csvSize;
        counter.docx += nextCallResult.docx;
        counter.docxSize += nextCallResult.docxSize;
        counter.xlsx += nextCallResult.xlsx;
        counter.xlsxSize += nextCallResult.xlsxSize;
    }

    return counter;
}
//This function looks within each message body preview in search for sharing invites, and if it is, then the function looks within the full message body to
//search for keywords (csv, docx, xlsx) so it can define what was shared.
async function emailSearch(call, user){
    let counter = {csv:0, docx:0, xlsx:0};
    for(let i = 0; i < call.value.length; i++){
        if(call.value[i].bodyPreview.includes('invited you to edit a file') || call.value[i].bodyPreview.includes('convidou você para editar um arquivo')){
            let messageWithHtml = await graphClient.api('users/' + `${user}` + '/messages/' + `${call.value[i].id}`).select('body').get();
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
        counter.csv += nextSearch.csv
        counter.docx += nextSearch.docx
        counter.xlsx += nextSearch.xlsx
    }
    return counter;
}

//**EMAIL LIST CALL PREVIEW
//let chamada = await graphClient.api('users/suporte02@grupounus.com.br/mailFolders/sentItems/messages').select('id,bodyPreview').top(1000).get();
//**ONEDRIVE LIST CALL PREVIEW
//let chamada = await graphClient.api('users/suporte02@grupounus.com.br/drive/root/children').top(100).select('file,folder,name,id,size').get();

let handle = await fs.open('C:/Users/fernando.garbato/Desktop/graph_demo/teste.txt', 'w')

async function officeSearch(usuarios) {
    for(let i = 0; i < usuarios.value.length; i++){
        console.log(i,' ', usuarios.value[i].displayName);
        if(usuarios.value[i].givenName !=null && usuarios.value[i].accountEnabled == true){
            //Defines the user and writes name
            let usuario = usuarios.value[i].userPrincipalName;
            await handle.writeFile(usuarios.value[i].displayName + ';');
            try{
                //Calls for email data about csv, docx and xlsx files shared, then writes that in csv format
                let emailCall = await graphClient.api('users/' + `${usuario}` + '/mailFolders/sentItems/messages').top(1000).select('id,bodyPreview').get();
                let emailInfo = await emailSearch(emailCall, usuario);
                await handle.writeFile(`${emailInfo.csv}` + ';' + `${emailInfo.xlsx}` + ';' + `${emailInfo.docx}` + ';');
            } catch(error){
                await handle.writeFile('\n');
                console.log(usuarios.value[i].displayName);
                console.log(error);
                continue;
            }
            try{
                //Calls for OneDrive file data, then writes that info in csv format
                let driveCall = await graphClient.api('users/' + `${usuario}` + '/drive/root/children').top(100).select('file,folder,name,id,size').get();
                let driveInfo = await folderSearch(usuario, driveCall);
                await handle.writeFile(`${driveInfo.xlsx}` + ';' + `${(driveInfo.xlsxSize/1024**2).toFixed(2)}` + ';' + `${driveInfo.csv}` + ';' + `${(driveInfo.csvSize/1024**2).toFixed(2)}` + ';' + `${driveInfo.docx}` + ';' + `${(driveInfo.docxSize/1024**2).toFixed(2)}`)
                await handle.writeFile('\n');
            } catch(error){
                await handle.writeFile('\n');
                console.log(usuarios.value[i].displayName);
                console.log(error);
                continue;

            }
        }
    }
}


//  let usersCall = await graphClient.api('users').select('userPrincipalName,displayName,givenName,accountEnabled').top(999).get();
//  await officeSearch(usersCall)


// let chamada = await graphClient.api('users/suporte03@grupounus.com.br/drive/root/children').top(100).select('file,folder,name,id,size').get();
// console.log(await folderSearch('suporte03@grupounus.com.br', chamada));

// let result = await graphClient.api('users/suporte02@grupounus.com.br/messages').get();
// const d = new Date("2022-03-25T00:00:00Z");
// console.log(result.value[0].receivedDateTime < d);


let emailUseHandle = await fs.open('C:/Users/fernando.garbato/Desktop/graph_demo/generated/email_use.txt', 'w')


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
        .filter("receivedDateTime ge 2024-10-13T00:00:00Z")
        .count(true)
        .get();
        let emailSentData = await graphClient.api("users/" + `${call.value[i].userPrincipalName}` + "/mailFolders/sentItems/messages")
        .select("id")
        .filter("sentDateTime ge 2024-10-13T00:00:00Z")
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

//   let emailUseReportCall = await graphClient.api("users").select("userPrincipalName, displayName, mail, surname, userType").get()
//   emailUseReport(emailUseReportCall);

console.log(await graphClient.api('users/suporte02@grupounus.com.br').select("userType").get())


//console.log(await graphClient.api("users/suporte02@grupounus.com.br/messages").get())

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