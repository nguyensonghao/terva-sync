/*
    Author: hao.nguyensong@powergatesoftware.com
    Date: Feb 14 2020
    Detail: get list file and folder in local and insert to database
*/

const fs = require('fs');
const _chunk = require('lodash').chunk;

const connection = require('./connection');
const DIR_FILE = "/home/nguyensonghao/PowerGate/Terva/terva-sync/Summit AG";
const ROOT_PATH_FILE = "Summit AG";

const getListFile = (dir, filelist, rootPath) => {
	files = fs.readdirSync(dir);
	filelist = filelist || [];
	files.forEach((file) => {
		let f = `${dir}/${file}`;		
    	if (fs.statSync(f).isDirectory()) {
            filelist.push({
                name: file,
                path: rootPath,
                isFolder: true,
                statusSync: 0
            })

    		let path = `${rootPath}/${file}`;
      		filelist = getListFile(f, filelist, path);
    	} else {
    		filelist.push({
    			name: file,
                path: rootPath,
                statusSync: 0
    		})
    	}
  	})

  	return filelist
}

const insertAllFileSubmitAg = async (list) => {
    let files = _chunk(list, 1000);
    for (let i = 0; i < files.length; i++) {
        await insertFile(files[i]);
    }
}

const insertFile = (list) => {
    return new Promise((resolve, reject) => {
        connection().then(db => {
            db.collection('submitAgPortalFiles').insertMany(list, (err, result) => {
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

(async() => {
    try {
        const listFile = getListFile(DIR_FILE, [], ROOT_PATH_FILE);
        await insertAllFileSubmitAg(listFile);
        console.log(listFile.length);
    } catch (err) {
        console.log(err);
    }
})();