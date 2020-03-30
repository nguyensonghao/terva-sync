/*
    Author: hao.nguyensong@powergatesoftware.com
    Date: Feb 14 2020
    Detail: get information for all file and folder from SP
*/

const request = require('request');
const config = require('config');

const connection = require('./connection');
const { parseResponse, getAccessToken, convertInformationFile, requestGet, generateBatchKey } = require('./helper');
const urlSP = config.get('SHAREPOINT.URL');
const ROOT_FOLDER_SP = config.get('SHAREPOINT.DOCUMENT_LIBRARY');
const GUID = config.get('SHAREPOINT.GUID')
const LIMIT_REQUEST = 100;

const getListFile = (skip) => {    
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('submitAgPortalFiles').find({
                // statusSync: 0
            }).skip(skip).limit(LIMIT_REQUEST).toArray((err, result) => {
                db.close();
                if (err) {
                    reject(err);
                } else {
                    resolve(result);
                }
            })
        })
    })    
}

const getNumberFile = () => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('submitAgPortalFiles').count({
                // statusSync: 0
            }, (err, size) => {
                db.close();
                if (err) {
                    reject(err);
                } else {
                    resolve(size);
                }                
            })
        })
    })   
}

const getMainMetadata = (accessToken) => {
    return new Promise((resolve, reject) => {
        const url = `${urlSP}/_api/Web/Lists(guid'${GUID}')/Fields?$filter=(Group eq 'Custom Columns') and (FromBaseType eq false) and (Hidden ne true)&$select=StaticName,FieldTypeKind`
        requestGet(accessToken, url).then(data => {
            let rs = data.results;
            let mainMetadata = ['Id'];
            rs.forEach(metadataInfo => {
                mainMetadata.push(metadataInfo.StaticName);
            })

            mainMetadata.push('ID');
            resolve(mainMetadata);
        }).catch(e => {
            reject(e);
        })
    })    
}

const createRequestGetInformation = (batchKey, file) => {
    const request = [];    
    request.push(`--${batchKey}`);
    request.push('Content-Type: application/http');
    request.push('Content-Transfer-Encoding: binary');
    request.push('');
    if (file.isFolder) {
        request.push(`GET ${urlSP}/_api/web/GetFolderByServerRelativePath(decodedurl='${ROOT_FOLDER_SP}/${encodeURIComponent(`${file.path}/${file.name}`)}')?$expand=ListItemAllFields HTTP/1.1`);
    } else {
        request.push(`GET ${urlSP}/_api/web/GetFileByServerRelativePath(decodedurl='/sites/SummitAgManagement/Properties/${ROOT_FOLDER_SP}/${encodeURIComponent(`${file.path}/${file.name}`)}')?$expand=ListItemAllFields HTTP/1.1`);
    }    

    request.push('Accept: application/json;odata=verbose');
    request.push('');
    return request;
}

const sendBatchRequest = (batchKey, accessToken, files) => {
    return new Promise((resolve, reject) => {
        const url = `${urlSP}/_api/$batch`;
        const headers = {
            credentials: "include",
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': `multipart/mixed; boundary=${batchKey}`,
            "Accept": "application/json;odata=verbose"
        }

        let body = [
            `Content-Type: multipart/mixed; boundary="${batchKey}`,
            'Content-Length: 592',
            'Content-Transfer-Encoding: binary',
            ''
        ]

        files.forEach(file => {
            body = body.concat(createRequestGetInformation(batchKey, file));
        })
            
        body.push('--' + batchKey + '--`');
        let stringBody = body.join('\r\n');
        request({
            headers: headers,
            url: url,
            body: stringBody,
            method: 'POST',
        }, (err, res, body) => {
            if (err) {
                reject(err);
            } else {
                resolve(parseResponse(body));
            }
        })
    })    
}

const updateInformation = (info) => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('submitAgPortalFiles').update({
                name: info.spInfo.Name,
                path: info.spInfo.PathReal
            }, {
                $set: {
                    statusSync: 1,
                    spInfo: info.spInfo,
                    meta: info.meta
                }
            }, (err) => {
                db.close();
                if (err) {
                    reject(err);
                } else {
                    resolve(true);
                }
            })
        })
    })
}

(async() => {
    try {
        const numberFile = await getNumberFile();
        let accessToken = await getAccessToken();
        let metadata = await getMainMetadata(accessToken);
        for (let i = 0; i < numberFile; i+= LIMIT_REQUEST) {
            console.log(`Number request ${i}`);
            let batchKey = generateBatchKey();
            let files = await getListFile(i);
            let accessToken = await getAccessToken();
            let results = await sendBatchRequest(batchKey, accessToken, files);
            for (let j = 0; j < results.length; j++) {
                let info = convertInformationFile(results[j], ROOT_FOLDER_SP, metadata);
                if (info) {
                    await updateInformation(info);
                }
            }
        }                
    } catch (err) {
        console.log(err);
    }
})();