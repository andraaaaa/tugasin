const express = require('express');
const path = require('path');
const fileURLToPath = require('url');
const fs = require('fs');
const bodyParser = require('body-parser');
const moment = require('moment');
const http = require('http');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const Blob = require('node-blob');
const FileSaver = require('file-saver');

const app = express();

app.use('/static', express.static('public'));
app.use(bodyParser.urlencoded({extended: false}));

app.get('/dashboard', function(req, res){
    res.sendFile(path.join(__dirname, './public/templates/dashboard.html'));
});

app.get('/pegawai', function(req, res){
    let json_data = fs.readFileSync('./public/json/pegawai.json');
    var datapeg = JSON.parse(json_data);
    console.log(datapeg); 

    res.sendFile(path.join(__dirname, './public/templates/pegawai.html'));
});

app.get('/generate', function(req, res){
    res.sendFile(path.join(__dirname, './public/templates/generatesurat.html'));
});

app.get('/setting', function(req, res){
    res.sendFile(path.join(__dirname, 'setting.html'));
});

app.get('/alokasipetugas', function(req, res){
    res.sendFile(path.join(__dirname, './public/templates/alokasipetugas.html'));
});

app.get('/backuprestore', function(req, res){
    res.sendFile(path.join(__dirname, './public/templates/generatesurat.html'));
});

app.get('/sandbox', function(req, res){
    res.sendFile(path.join(__dirname, 'sandbox.html'));

    let data_adds = {
        pangkatgolongan: null,
        nama_seksi: null
    };

    let jpegawai = fs.readFileSync('./public/json/pegawai.json');
    var jp = JSON.parse(jpegawai);
    

});

app.get('/addmitra', function(req, res){
    res.sendFile(path.join(__dirname, 'tambahmitra.html'));
});

app.post('/create-mitra', function(req, res){
    let json_data = fs.readFileSync('./public/json/mitra.json');
    var datamitra = JSON.parse(json_data);
    var ap = req.body.desa + ' ' + req.body.kec;
    try {
        var mitra = {
            nama: req.body.nama,
            asal_petugas: ap,
            golongan: "-",
            ruang: "-",
            jabatan: "Mitra Statistik",
            nip: "-",
            jk: req.body.jk,
            kec: req.body.kec,
            desa: req.body.desa,
            alamat: req.body.alamat,
            ktp: req.body.ktp,
            nohp: req.body.nohp,
            email: req.body.email,
            bank: req.body.bank,
            no_rek: req.body.norek,
            no_npwp: req.body.npwp

        };

        datamitra['mitra'].push(mitra);
        var mitra_str = JSON.stringify(datamitra);

        fs.writeFileSync('./public/json/mitra.json', mitra_str, function(err){
            if(err){
                console.log("Tambah mitra gagal. Coba lagi");
            } else {
                console.log("Tambah mitra berhasil : " + req.body.nama)
            }
        });
    } catch (error){
        console.log(error);
    }

    res.redirect('/pegawai');
});

app.post('/create-kegiatan', function(req, res){
    let json_data = fs.readFileSync('./public/json/kegiatan.json');
    var json_kegiatan = JSON.parse(json_data);
    
    
})

app.post('/send-data', function(req, res){
    let json_data = fs.readFileSync('./public/json/ganttsurat.json');
    let json_pegawai = fs.readFileSync('./public/json/pegawai.json');
    let configjson = fs.readFileSync('./public/json/suratconfig.json');
    var datapeg = JSON.parse(json_pegawai);
    var datasurat = JSON.parse(json_data);
    var conf = JSON.parse(configjson);
    var nomor_surat_fix = parseInt(conf['init_surat']) + datasurat.length;
    var namasurat = req.body.nama
    var vgol, vjab, vnip, vpgkt, vjabkf, vasal, vks;

    for(var i=0; i < datapeg.length; i++){
        if(datapeg[i].nama == namasurat){
            vnip = datapeg[i].nip;
            vgol = datapeg[i].golongan;
            vjab = datapeg[i].jabatan;
            vpgkt = datapeg[i].ruang;
            vasal = datapeg[i].asal_petugas;
        } else continue;
    }

    for (var j=0; j<datapeg.length; j++){
        if(datapeg[j].nama == req.body.kfsender){
            vjabkf = datapeg[j].jabatan;
            vks = datapeg[j].kode_seksi;
        } else continue;
    }

    var ts = new Date();    
    var bln = ts.getMonth()+1;
    var ts_complete = ts.getDate() + '-' + bln + '-' + ts.getFullYear(); + ' ' + ts.getHours() + ':' + ts.getMinutes() + ':' + ts.getSeconds()
    var ns = "B." + nomor_surat_fix + "/6401" + vks + "/" + req.body.klas_desk + "."+ req.body.klas_kode + "/" + bln + "/" + thn;
    var t_st = moment(req.body.tanggalst, 'MM/DD/YYYY', true).locale("id").format("LL");
    var tp = "";
    try {
        var sdf = moment(req.body.startdate, 'MM/DD/YYYY', true).locale("id").format("LL");
        var fdf = moment(req.body.finishdate, 'MM/DD/YYYY', true).locale("id").format("LL");
        if(fdf === "Invalid date"){
            fdf = sdf;
            tp = sdf;
        } else {
            tp = sdf + " s.d. " + fdf;
        }

        var dc = getDateCount(sdf, fdf);
        dc = Math.abs(1 + dc/86400000);
        if(dc === 1.0000000115740741){
            dc = 1;
        }
    
    //var outst = "surattugas_" + getSkrg('_') + '_' + t.getHours() + t.getMinutes() + t.getSeconds() + ".docx";
        var ganttdata = {
                timestamp: ts_complete,
                tipe: req.body.tipe,
                tipe_surat: req.body.tipe_surat,
                kodesatker: vks,
                kf_satker: req.body.kfsender,
                jabatan_kf: vjabkf,
                tanggal_ST: t_st,
                nomorsurat: nomor_surat_fix,
                nomorsuratfull: ns,
                bulan: bln,
                tahun: thn,
                nama: req.body.nama,
                nip: vnip,
                pangkat: vpgkt,
                golongan: vgol,
                jabatan: vjab,
                asal_petugas: vasal,
                tujuan: req.body.tujuan,
                kegiatan: req.body.keg,
                tanggal_pelaksanaan: tp,
                start_date: sdf,
                finish_date: fdf,
                date_count: dc
        };

        datasurat.push(ganttdata);
        var surat_p = JSON.stringify(datasurat);

        fs.writeFileSync('./public/json/ganttsurat.json', surat_p, function(err){
            if(err){
                console.log("Tambah data gagal. Coba lagi");
            } else {
                console.log("Tambah data berhasil");
            }
        });
        } catch (error){
            console.log(error);
        }
        console.log(datasurat);
        res.redirect('/dashboard');
});

app.post('/ubah-settings', function(req, res){
    let configjson = fs.readFileSync('./public/json/suratconfig.json');
    var cj = JSON.parse(configjson);

    cj.kepalabps =  req.body.kepalabps;
    cj.nip_kepala = req.body.nipkepala;
    cj.ppk = req.body.ppk;
    cj.nip_ppk = req.body.nipppk;
    cj.init_surat = req.body.initsurat;

    var cj_str = JSON.stringify(cj);
    fs.writeFileSync('./public/json/suratconfig.json', cj_str, function(err){
        if(err){
            console.log("Update data gagal. Coba lagi");
        } else {
            console.log("Berhasil!");
        }
    });
    res.redirect('/dashboard'); 
});

app.get('/unduh', function(req, res){
    var dt, f;
    console.log(__dirname);
    var nomor = req.query.nomor;
    var fileku = fs.readFileSync('./public/json/ganttsurat.json');
    var myjson = JSON.parse(fileku);

    for(var i=0; i < myjson.length; i++){
        if (myjson[i].nomorsurat == nomor){
            dt = myjson[i];
        } else { continue; }
    }
    console.log(dt);
    var jenis_surat = dt.tipe_surat;
    if(jenis_surat === 'kunjungan'){
        f = "public/surat/surat_kunjungan.docx";
    } else if(jenis_surat === 'spd'){
        f = "public/surat/spd.docx"
    }

    generateSurtug(dt, f);
    var nms = nomor + ".docx";
    const directoryPath = __dirname + "/public/surat/generated/";
    res.download(directoryPath + nms, nms, (err) => {
        if (err) {
        res.status(500).send({
            message: "Could not download the file. " + err,
        });
        }
    });
});

app.get('/rekap', function(req, res){
    res.sendFile(path.join(__dirname, 'rekapitulasi.html'));    
});

function replaceErrors(key, value) {
    if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
            error[key] = value[key];
            return error;
        }, {});
    }
    return value;
}

function errorHandler(error) {
    console.log(JSON.stringify({error: error}, replaceErrors));

    if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map(function (error) {
            return error.properties.explanation;
        }).join("\n");
        console.log('errorMessages', errorMessages);
    }
    throw error;
}

function generateSurtug(s, f){
    var content = fs.readFileSync(path.resolve(__dirname, f), 'binary');
    var zip = new PizZip(content);
    var doc;
    try {
        doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    } catch(error) {
        errorHandler(error);
    }
    doc.setData(s);
    try {
        doc.render();
    }
    catch(error) {
        errorHandler(error);
    }
    var buf = doc.getZip().generate({
        type:"nodebuffer",
    });
    var namasurat = s.nomorsurat + '.docx';

    fs.writeFileSync(path.resolve(__dirname, 'public/surat/generated/'+namasurat), buf);
    /*var docblob = new Blob([buf], 
        {type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

    FileSaver.saveAs(docblob, namasurat);*/
}


function getDateCount(st, fin){
    var diffres;
    var d1 = new Date(st);
    var d2 = new Date(fin);
    var diff = Math.abs(d2-d1);
    if(diff === 0){
      diffres = 1
    } else diffres = diff;
    return diffres;
  }

var server = http.createServer(app);
app.listen(8080, '127.0.0.1', function() {
    server.close(function(){
        server.listen(3000, '127.0.0.1');
        console.log('Server running!');
    });
});