const express = require('express');
const bodyParser = require('body-parser')
require('dotenv').config();
const { Looker40SDK } = require('@looker/sdk');
const { NodeSession, NodeSettings} = require('@looker/sdk-node');
const ExcelJS = require('exceljs')
const nodemailer = require('nodemailer')

const ns = new NodeSettings('',{base_url:process.env.LOOKER_HOST});
ns.readConfig = () => {
    return(
        {client_id:process.env.LOOKER_CLIENT_ID, client_secret:process.env.LOOKER_CLIENT_SECRET}
    )
}
const session = new NodeSession(ns);

const sdk = new Looker40SDK(session);

const app = express()

app.use(bodyParser.json())

const credentials = {
    host:process.env.SMTP_ADDRESS,
    port:process.env.SMTP_PORT,
    secure:true,
    auth: {
        user:process.env.SMTP_USERNAME,
        pass:process.env.SMTP_PASSWORD
    },
    tls : { rejectUnauthorized: false }
}

console.log(credentials)

let _email = nodemailer.createTransport(credentials);

const formatFilters = (selectedFilters, dashboardFilters) => {
    try {
        let _filters = {}
        Object.entries(selectedFilters)?.map(([key,value]) => {
            let match = dashboardFilters?.find(f => (f.title === key));
            if (match) {
                _filters[match.dimension] = value
            }
        })
        return _filters
    } catch (ex) {
        console.error(`Error formatting filters: ${ex}`)
        return {}
    }
}
// const formatFilters = (dashboard_filters) => {
//     let _filters = {}
//     dashboard_filters?.map((filter) => {
//         _filters[filter.dimension] = filter.default_value
//     })
//     console.log(_filters)
//     return _filters
// }

const downloadData = async (query_id, format, tableCalcs) => {
    try {
        let _limit = -1;
        if (query_id) {
            if (tableCalcs) _limit = 100000
            const res = await sdk.ok(sdk.run_query({query_id:query_id.toString(), result_format:format, limit:_limit, apply_vis:true, apply_formatting:true}));
            return res;
        }
    } catch(ex) {
        console.error(`Error downloading data: ${ex}`)
    }

    return ""
}

const createExcelFunc = async (id, selectedFilters) => {
    if (!id) {
        return ''
    }          
    let {title, dashboard_elements, dashboard_filters} = await sdk.ok(sdk.dashboard(id, 'title,dashboard_elements,dashboard_filters'))
    let {blob} = await createBlob(dashboard_elements, dashboard_filters, selectedFilters)
    return {blob:blob, title:title}
}

const createBlob = async (elements, dashboard_filters, selectedFilters) => {
    let _filters = formatFilters(selectedFilters, dashboard_filters)
    let dataTiles = elements?.filter(tile => tile.query_id != null)

    let workbook = new ExcelJS.Workbook();

    for await (let tile of dataTiles) {
        let {query} =  tile;
        let query_id = tile.query_id
        if (query) {
            delete query.client_id
            delete query.can
            delete query.id
            query.filters = _filters || {}
            let {id} = await sdk.ok(sdk.create_query(query))
            if (id) {
                query_id = id;
            }
        let data = await downloadData(query_id, 'json', tile.query?.has_table_calculations);
        let worksheet = workbook.addWorksheet(tile.title.replace(/[^a-zA-Z0-9 ]/g, ' '))
        if (data.length > 0) {
            worksheet.columns = Object.keys(data[0])?.map(row => {return {header: row, key:row }})
            await worksheet.addRows(data);
        }
        }
    }
    var buff = await workbook.xlsx.writeBuffer().then(function (data) {
        var blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
        return blob;
    });
        
    return {'blob':buff}
}

app.get('/', (req,res) => {
    res.send("Service is running")
})

app.post('/', (req,res) => {
    let response = {'integrations':[{
        'name':'test_action',
        'label':'Download Excel Without Limit',
        'description':'This action allows you to send a dashboard and recieve an Excel file in return',
        'supported_action_types':['dashboard'],
        'params':[],
        'url':process.env.GCP_FUNCTION_EXECUTE_URL,
        'form_url':process.env.GCP_FUNCTION_FORM_URL
    }]}
    res.send(response)
})

app.post('/execute', async (req,res) => {
    console.log(req.body)
    let {url} = req.body?.scheduled_plan;
    let newURL = new URL(url);
    let dashboardId = newURL.pathname.replace("/dashboards/",'')
    console.log("search",newURL.search)
    let selectedFilters = {}
    let _selectedFilters = decodeURIComponent(newURL.search).replace("?",'')
    if (_selectedFilters != "") {
        _selectedFilters.split("&").map(row => {
            let array = row.split("=");
            selectedFilters[array[0].replace(/\+/g,' ')] = array[1].replace(/\+/g,' ')
        })
    }
    console.log("path",newURL.pathname)
    console.log("filter array",selectedFilters)
    
    let {emails} = req.body?.form_params;
    let {blob,title} = await createExcelFunc(dashboardId, selectedFilters)
    
    let message = {
        from:`${process.env.SMTP_EMAIL_FROM}`,
        to:`${emails}`,
        subject:`${title.replace(/[^a-zA-Z0-9 ]/g, ' ')}`,
        attachments: [
            {
                filename: `${title}.xlsx`,
                content: await Buffer.from(await blob.arrayBuffer())
            }
        ]
    }

    try {
        _email.sendMail(message, error => {
            if (error) {
                console.log("error", error)
            } else {
                res.json({status:'Message Sent'})
            }
        })
    } catch (ex) {
        console.error(`Error sending mail: ${ex}`)
    }

    res.send({success: true, message:req.body})
})

app.post('/form', (req,res) => {
    let fields = {
        'fields':[{
            name: "emails",
            label: "Email Addresses",
            type: "string",
            required: true,
        }
    ]}
    res.send(fields)
})

app.post('/createExcel', async (req,res) => {
    let id = req.body['dashboardId'];
    let _return = await createExcelFunc(id)
    
    let message = {
        from:'aaron.modic@bytecode.io',
        to:'aamodic@gmail.com',
        subject:'Test',
        attachments: [
            {
                filename: 'test.xlsx',
                content: await Buffer.from(await _return.arrayBuffer())
            }
        ]
    }

    _email.sendMail(message, error => {
        if (error) {
            console.log("error", error)
        } else {
            res.json({status:'Message Sent'})
        }
    })
    res.send(_return)
})

//app.listen(8000, () => console.log(`Server started on 8000`))

exports.api = app