/*
    Author: hao.nguyensong@powergatesoftware.com
    Date: Feb 14 2020
    Detail: convert information of file to format store file in database
*/

const config = require('config');

const connection = require('./connection');
const PORTAL_ID = config.get('SHAREPOINT.PORTAL_ID');

const getListFile = () => {    
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('submitAgPortalFiles').find({
                statusSync: 1
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

const insertFile = (file) => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('submitAgPortalFileFullInfo').insert(file, (err, result) => {
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

const addInformation = (file) => {
    return {
        name: file.name,
        path: `${PORTAL_ID}/${file.path}`,
        portalId: PORTAL_ID,
        syncSP: true,
        meta: file.meta,
        spInfo: file.spInfo,
        isFolder: file.isFolder,
        createdAt: new Date(),
        updatedAt: new Date()        
    }
}

(async() => {
    try {
        const files = await getListFile();
        for (let i = 0; i < files.length; i++) {
            await insertFile(addInformation(files[i]));
        }
    } catch (err) {
        console.log(err);
    }
})();