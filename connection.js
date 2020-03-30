const MongoClient = require('mongodb').MongoClient;

module.exports = () => {
	return new Promise((resolve, reject) => {
		MongoClient.connect('mongodb://localhost:27017/terva', (err, db) => {
            if (err) {
                reject(err);
            } else {
                resolve(db);
            }	        
	    })
	})
}