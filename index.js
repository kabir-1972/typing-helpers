const http = require('http');
const url = require('url');
const Converter = require('./subConverter');
const converter = new Converter();

const server = http.createServer((req, res) => {
    if (req.method === 'POST') {
        let body = '';
        req.setEncoding('utf8');
        req.on('data', chunk => { body += chunk; });
        req.on('end', () => {
            const result = converter.Convert(body);
            //let result = http.responseText;

            const FindText = "h" + String.fromCharCode(0x09BC);
            const findText2 = "h" + String.fromCharCode(0x2021) + String.fromCharCode(0x09BC);

            const findText2r = "W" + String.fromCharCode(0x09BC);
            const findText2r2 = "W" + "w" + String.fromCharCode(0x09BC);

            const findText2dr = "W" + "w" + String.fromCharCode(0x09BC);

            // Replace occurrences
            result = result.replaceAll(findText2, "‡q");
            result = result.replaceAll(findText2r2, "wo");
            result = result.replaceAll(FindText, "q");
            result = result.replaceAll(findText2r, "o");

            res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
            res.end(result);
        });

    } else if (req.method === 'GET') {
        const queryObject = url.parse(req.url, true).query;
        
        if (queryObject.bangla) {
            // Decode UTF-8 Bangla input
            const inputText = decodeURIComponent(queryObject.bangla);
            const result = converter.Convert(inputText);

            res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
            res.end(result);
        } else {
            res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
            res.end("No 'bangla' parameter found");
        }

    } else {
        res.writeHead(405, { "Content-Type": "text/plain; charset=utf-8" });
        res.end("Only GET and POST methods are allowed");
    }
});

server.listen(1337, () => console.log("Server running on port 1337"));
