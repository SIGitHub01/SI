var fs = require('fs');
var htmlparser = require("htmlparser2");
const { BlobServiceClient } = require('@azure/storage-blob');
const { v1: uuid} = require('uuid');
const appInsights = require('applicationinsights');
appInsights.setup().start();
appInsights.defaultClient.config.samplingPercentage = 100; // 33% of all telemetry will be sent to Application Insights
appInsights.start();

context.log('I was here **********');

function traverse(an_array) {

    var rval = "";

    an_array.forEach(function (element) {
        if (element.type == "tag" && element.name == "b") {
            var the_value = element.children[0].data;
            if (the_value == "Partner Name") {
                var first_tr = element.parent.parent.parent.parent;
                var count = 0;
                first_tr.children.forEach(function (element2) {
                    if (element2.type = "tag") {
                        if (element2.name == "tr" && count > 0) {

                            var partner_name = element2.children[1].children[1].children[0].data;
                            partner_name = partner_name.trim();

                            if (partner_name.length > 0) {
                                var last_file_name = element2.children[3].children[1].children[0].data;
                                last_file_name = last_file_name.trim();

                                var recvd_date = element2.children[5].children[1].children[0].data;
                                recvd_date = recvd_date.trim();

                                rval = rval
                                    + '"' + partner_name + '", '
                                    + '"' + last_file_name + '", '
                                    + '"' + recvd_date + '"\n';
                            }

                        } else {
                            count++;
                        }
                    }
                });
            }
        }
        if ( element.children != null)
        {
            rval = rval + traverse(element.children);
        }
    });
    return rval;
}

function parse_mail_for_partners(the_mail_string) {

    var the_dom = null;

    var handler = new htmlparser.DefaultHandler(function (error, dom) {
        if (error)
            context.log(error);
        else {
            the_dom = dom;
        }
    });
    var parser = new htmlparser.Parser(handler);
    parser.parseComplete(the_mail_string);

    return traverse(the_dom)
}

// A helper function used to read a Node.js readable stream into a string
async function streamToString(readableStream) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        readableStream.on("data", (data) => {
            chunks.push(data.toString());
        });
        readableStream.on("end", () => {
            resolve(chunks.join(""));
        });
        readableStream.on("error", reject);
    });
}

const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING;
const MAIL_CONTAINER_NAME = process.env.MAIL_CONTAINER_NAME;
const MAIL_INBOUND_BLOB_NAME = process.env.MAIL_INBOUND_BLOB_NAME;
const MAIL_OUTBOUND_BLOB_NAME = process.env.MAIL_OUTBOUND_BLOB_NAME;

async function main() {

    context.log('start');

    const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);

    // get the data
    context.log('start getContainerClient ' + MAIL_CONTAINER_NAME);
    const containerClient = blobServiceClient.getContainerClient(MAIL_CONTAINER_NAME);
    context.log('stop getContainerClient ' + MAIL_CONTAINER_NAME);

    context.log('start getBlockBlobClient ' + MAIL_INBOUND_BLOB_NAME);
    const inboundBlockBlobClient = containerClient.getBlockBlobClient(MAIL_INBOUND_BLOB_NAME);
    const downloadBlockBlobResponse = await inboundBlockBlobClient.download(0);
    context.log('start getBlockBlobClient ' + MAIL_INBOUND_BLOB_NAME);

    context.log('start streamToString ');
    var inbound_data_body = await streamToString(downloadBlockBlobResponse.readableStreamBody)
    context.log('finish streamToString ');

    // parse the data
    context.log('start parse ');
    var outbound_data_body = parse_mail_for_partners(inbound_data_body);
    context.log('stop parse ');

    // write the data
    context.log('start write ' + MAIL_OUTBOUND_BLOB_NAME);
    const outboundBlockBlobClient = containerClient.getBlockBlobClient(MAIL_OUTBOUND_BLOB_NAME);
    const uploadBlobResponse = await outboundBlockBlobClient.upload(outbound_data_body, outbound_data_body.length);
    context.log('stop write ' + MAIL_OUTBOUND_BLOB_NAME);

}

main().then(() => context.log('Done')).catch((ex) => context.log(ex.message));