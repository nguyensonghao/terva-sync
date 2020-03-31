const request = require('request');
const config = require('config');
const fs = require('fs');

const SP_URL = config.get('SHAREPOINT.URL')
const TENANTID = config.get('SHAREPOINT.TENANTID');
const CLIENT_ID = config.get('SHAREPOINT.CLIENT_ID');
const CLIENT_SECRET = config.get('SHAREPOINT.CLIENT_SECRET');
const RESOURCE = config.get('SHAREPOINT.RESOURCE');
const DOCUMENT_LIBRARY = config.get('SHAREPOINT.DOCUMENT_LIBRARY')

const grep = (elems, callback, invert) => {
    let callbackInverse,
        matches = [],
        i = 0,
        length = elems.length,
        callbackExpect = !invert;
    // Go through the array, only saving the items  
    // that pass the validator function  
    for (; i < length; i++) {
        callbackInverse = !callback(elems[i], i);
        if (callbackInverse !== callbackExpect) {
            matches.push(elems[i]);
        }
    }

    return matches;
}

const parseResponse = (batchResponse) => {
    let results = grep(batchResponse.split("\r\n"), function (responseLine) {
        try {
            return responseLine.indexOf("{") != -1 && typeof JSON.parse(responseLine) == "object";
        } catch (ex) {
            /*adding the try catch loop for edge cases where the line contains a { but is not a JSON object*/
        }
    }, null);
    
    return results.map(function (result) {
        return JSON.parse(result);
    })
}

const getAccessToken = () => {
    return new Promise((resolve, reject) => {
        const options = {
            method: 'POST',
            url: `https://accounts.accesscontrol.windows.net/${TENANTID}/tokens/OAuth/2`,
            headers: {
                'content-type': 'application/x-www-form-urlencoded'
            },
            form: {
                grant_type: 'client_credentials',
                resource: RESOURCE,
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET
            }
        }

        request(options, (error, response, body) => {
            if (error) {
                reject(error);
            } else {
                let data = JSON.parse(body);
                if (data.error) {
                    reject(new Error(data.error_description));
                } else {
                    resolve(data.access_token);
                }
            }
        })
    })
}

const getPath = (fullPath) => {
    let pathArr = fullPath.split('/');
    if (pathArr.length) {
        pathArr.pop();
        return pathArr.join('/');
    }   
    
    return fullPath;
}

const convertMetadata = (metadata, mainMetaData) => {
    let result = {};
    mainMetaData.forEach(item => {
        result[item] = metadata[item];
    })

    return result;
}

const convertInformationFile = (result, ROOT_FOLDER_SP, mainMetaData) => {
    if (result && result.d) {
        let info = result.d;
        let path = info.ServerRelativeUrl.replace(`/sites/SummitAgManagement/Properties/${ROOT_FOLDER_SP}/`, '');
        return {
            spInfo: {
                UniqueId: info.UniqueId,
                ItemCount: info.ItemCount,
                Name: info.Name,
                TimeCreated: info.TimeCreated,
                TimeLastModified: info.TimeLastModified,
                ServerRelativeUrl: info.ServerRelativeUrl,
                Path: path,
                PathReal: getPath(path)
            },
            meta: convertMetadata(info.ListItemAllFields, mainMetaData)
        }
    } else {
        return null;
    }    
}

const createFolder = (path, id, accessToken) => {
    return new Promise((resolve, reject) => {
        try {
            const options = {
                method: 'POST',
                url: `${SP_URL}/_api/web/GetFolderByServerRelativeUrl('Land Files')/folders?$expand=ListItemAllFields`,
                headers: {
                    'authorization': `Bearer ${accessToken}`,
                    'content-type': 'application/json;odata=verbose',
                    'accept': 'application/json;odata=verbose'
                },
                body: JSON.stringify({
                    "__metadata": {
                        "type": "SP.Folder"
                    },
                    "ServerRelativeUrl": `${DOCUMENT_LIBRARY}/${path}`
                })
            }
            
            request(options, function (error, response, body) {
                try {
                    if (error) {
                        reject(error);
                    } else {
                        let data = JSON.parse(body);
                        if (data.error) {
                            resolve({
                                status: false,
                                id: id
                            })
                        } else {
                            if (data.error_description) {
                                console.log(data.error_description);
                                resolve({
                                    status: false,
                                    id: id
                                })
                            } else {
                                resolve({
                                    status: true,
                                    id: id
                                })
                            }
                        }
                    }   
                } catch (error) {
                    console.log(error.message);
                    resolve({
                        status: false,
                        id: id
                    })
                }            
            })   
        } catch (error) {
            resolve({
                status: false,
                id: id
            })
        }        
    })
}

const uploadFile = (localPath, folderPath, fileName, id, accessToken) => {
    return new Promise((resolve, reject) => {
        try {
            let url = `${SP_URL}/_api/web/GetFolderByServerRelativeUrl('${DOCUMENT_LIBRARY}/${folderPath}')/Files/add(url='${fileName}',overwrite=true)?$expand=ListItemAllFields`;
            if (!folderPath) {
                url = `${SP_URL}/_api/web/GetFolderByServerRelativeUrl('${DOCUMENT_LIBRARY}/')/Files/add(url='${fileName}',overwrite=true)?$expand=ListItemAllFields`;
            }
            
            fs.readFile(localPath, (err, data) => {
                const options = {
                    method: 'POST',
                    url: url,
                    headers: {            
                        'authorization': `Bearer ${accessToken}`,
                        'content-type': 'application/json;odata=verbose',
                        'accept': 'application/json;odata=verbose'
                    },
                    body: data
                }
            
                request(options, (error, response, body) => {
                    if (error) {
                        reject(error);
                    } else {
                        if (!body) {
                            resolve(true);
                        } else {
                            try {
                                let data = JSON.parse(body);
                                if (data.error) {
                                    resolve({
                                        status: false,
                                        id: id
                                    })
                                } else {
                                    if (data.error_description) {
                                        console.log(data.error_description);
                                        resolve({
                                            status: false,
                                            id: id
                                        })
                                    } else {
                                        resolve({
                                            status: true,
                                            id: id
                                        })
                                    }
                                }   
                            } catch (error) {
                                console.log(error.message);
                                resolve({
                                    status: false,
                                    id: id
                                })
                            }                            
                        }   
                    }                                     
                })
            })     
        } catch (error) {
            resolve({
                status: false,
                id: id
            })
        }          
    })    
}

const requestGet = (accessToken, url) => {
    return new Promise((resolve, reject) => {
        const options = {
            method: 'GET',
            url,
            headers: {
                'authorization': `Bearer ${accessToken}`,
                'content-type': 'application/json;odata=verbose',
                'accept': 'application/json;odata=verbose'
            }
        }

        request(options, function (error, response, body) {
            try {
                if (error) {
                    reject(error);
                } else {
                    let data = JSON.parse(body);
                    if (data.error) {
                        reject(new Error(data.error.message.value));
                    } else {
                        if (data.error_description) {
                            reject(new Error(data.error_description));
                        } else {
                            resolve(data.d);
                        }
                    }
                }   
            } catch (error) {
                reject(error);
            }            
        })
    })    
}

const makeid = (length) => {
    let result = '';
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    const charactersLength = characters.length;
    for (let i = 0; i < length; i++) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }

    return result;
}

const generateBatchKey = () => {
    let key = `batch__${makeid(11)}`;
    return key;
}

module.exports = {
    generateBatchKey,
    parseResponse,
    getAccessToken,
    requestGet,
    convertInformationFile,
    createFolder,
    uploadFile
}