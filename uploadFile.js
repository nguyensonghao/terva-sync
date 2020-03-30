/*
    Author: hao.nguyensong@powergatesoftware.com
    Date: Feb 18 2020
    Detail: get list file and folder in local and insert to database
*/

/*
    * Note:
        - We have about 100,000 files and folders to upload Sharepoint
        - Sharepoint does not support uploading folder (only support upload file and create folder) and not support to send batch request for upload file
    * Steps to upload file:
        - Get information of all files and folders to save in database including name, path, level. For example: we have folder "Upload" and it includes file "Upload1.txt" and folder "Upload1". "Upload" has level 0. "Upload1.txt" and "Upload1" have level 1. Add "level" attribute to help us create folders and upload files to SharePoint in the correct order. (See more in my code)
        - At the same time, we will create multi request to SP to upload file(NUMBER_REQUEST_PARALLEL variable). It will reduce the upload time.
    
    => Because we have many files i think we should create a new Sharepoint to test this script first to make sure it does not affect Live SharePoint.
*/

const fs = require('fs');
const _chunk = require('lodash').chunk;
const ObjectID = require('mongodb').ObjectID;

const connection = require('./connection');
const Helper = require('./helper');
const FOLDER_UPLOAD_NAME = "UploadTest";
const DIR_FILE = `${__dirname}/${FOLDER_UPLOAD_NAME}`;
const NUMBER_REQUEST_PARALLEL = 5;

const getListFile = (dir, filelist, rootPath, level) => {
	files = fs.readdirSync(dir);
	filelist = filelist || [];
	files.forEach((file) => {
		let f = `${dir}/${file}`;		
    	if (fs.statSync(f).isDirectory()) {
            filelist.push({
                name: file,
                path: rootPath,
                isFolder: true,
                statusSync: 0,
                level: level
            })
            
            let path = rootPath ? `${rootPath}/${file}` : file;
      		filelist = getListFile(f, filelist, path, level + 1);
    	} else {
    		filelist.push({
    			name: file,
                path: rootPath,
                statusSync: 0,
                level: level
    		})
    	}
  	})

  	return filelist
}

const sendRequestParallel = (listFile, accessToken) => {
    return new Promise((resolve, reject) => {
        let listPromise = [];
        listFile.forEach(file => {
            if (file) {
                if (file.isFolder) {
                    if (file.path) {
                        listPromise.push(Helper.createFolder(`${file.path}/${file.name}`, file._id, accessToken))
                    } else {
                        listPromise.push(Helper.createFolder(`${file.name}`, file._id, accessToken))
                    }                    
                } else {
                    if (file.path) {
                        listPromise.push(Helper.uploadFile(`./${FOLDER_UPLOAD_NAME}/${file.path}/${file.name}`, file.path, file.name, file._id, accessToken));
                    } else {
                        listPromise.push(Helper.uploadFile(`./${FOLDER_UPLOAD_NAME}/${file.name}`, file.path, file.name, file._id, accessToken));
                    }
                }
            }
        })

        Promise.all(listPromise).then(result => {
            resolve(result);
        })
    })
}

const insertToDB = (listFile) => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('fileUploads').insertMany(listFile, (err, result) => {
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

const insertAllFile = async (list) => {
    let files = _chunk(list, 1000);
    for (let i = 0; i < files.length; i++) {
        await insertToDB(files[i]);
    }
}

const getListLevel = () => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('fileUploads').distinct('level', {
                statusUpload: {
                    $ne: 1
                }
            }, (err, list) => {
                db.close();
                if (err) {
                    reject(err);
                } else {
                    resolve(list);
                }
            })
        })
    })
}

const getListFileByLevel = (level) => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('fileUploads').find({
                level, 
                statusUpload: {
                    $ne: 1
                }
            }).toArray((err, result) => {
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


const updateStatusUpload = (item) => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('fileUploads').update({
                _id: new ObjectID(item.id)
            }, {
                $set: {
                    statusUpload: 1
                }
            }, (err) => {
                db.close();
                resolve(err ? false : true);
            })           
        })
    })
}

(async() => {
    try {
        let listFile = getListFile(DIR_FILE, [], null, 0);
        await insertAllFile(listFile);

        let listLevel = await getListLevel();        
        listLevel.sort(function(a, b){ return a-b });
        for (let i = 0; i < listLevel.length; i++) {
            let listFileByLevel = await getListFileByLevel(listLevel[i]);
            if (listFileByLevel.length) {
                for (let j = 0; j < listFileByLevel.length; j+=NUMBER_REQUEST_PARALLEL) {
                    console.log(j);
                    let accessToken = await Helper.getAccessToken();
                    let listTemp = [], size = j + NUMBER_REQUEST_PARALLEL < listFileByLevel.length ? j + NUMBER_REQUEST_PARALLEL : listFileByLevel.length;
                    for (let k = j; k < size; k++) {
                        listTemp.push(listFileByLevel[k]);
                    }

                    let result = await sendRequestParallel(listTemp, accessToken);
                    result.forEach(async item => {
                        if (item.status) {
                            await updateStatusUpload(item);
                        }
                    })
                }
            }
        }
    } catch (err) {
        console.log(err);
    }
})();