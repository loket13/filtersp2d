const jenisSPM = ['SEMUA','NON GAJI', 'NON GAJI KONTRAKTUAL', 'GUP', 'TUNJANGAN KINERJA SUSULAN', 'GAJI LAINNYA', 'GAJI INDUK', 'SPM-KP-PAJAK', 'KEKURANGAN GAJI', 'GUP KKP', 'PENGHASILAN PPNPN INDUK', 'PEMBAYARAN RPATA', 'GTUP NIHIL', 'TUNJANGAN KINERJA BULANAN', 'SPM THR TUNKIN', 'GAJI LAINNYA PPPK', 'GAJI SUSULAN', 'PENGHASILAN PPNPN SUSULAN', 'SPM THR GAJI PNS/TNI/POLRI', 'SPM GAJI 13 PNS/TNI/POLRI', 'GAJI PPPK INDUK', 'KEKURANGAN TUNJANGAN KINERJA', 'SPM GAJI 13 TUNKIN', 'SPM THR PPNPN', 'RETUR', 'UP', 'GAJI SUSULAN PPPK', 'SPM-P-PNBP', 'SPM THR PPPK', 'SPM GAJI 13 PPNPN', 'TUP', 'SPM GAJI 13 PPPK', 'KEKURANGAN GAJI PPPK', 'GUP VALAS', 'PENIHILAN RPATA', 'SPM-IB-PAJAK', 'GUP NIHIL', 'PENGESAHAN BLU', 'GUP VALAS NIHIL', 'PENGESAHAN HIBAH', 'SP4HL', 'SPM GAJI 13 PEJABAT NEGARA', 'SPM THR PEJABAT NEGARA'];

const bulanSPM = ['SEMUA', '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
const akunSPM = ['SEMUA', '51', '52', '53']


function upload() {
    var files = document.getElementById('file_upload').files;
    if(files.length==0){
      alert("Pilih file yang sesuai...");
      return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    }else{
        alert("Please select a valid excel file.");
    }
  }
   
function excelFileToJSON(file){
      try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {
   
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type : 'binary'
            });
            var result = {};
            workbook.SheetNames.forEach(function(sheetName) {
                var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result['data'] = roa;
                }
            });
            tesLoop(result);
            }
        }catch(e){
            console.error(e);
        }
};

  async function s2d(str){
    mma = ["null","JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
    raw = str.split('-');
    ndd = raw[0].toString();
    nmr = raw[1]; 
    nyy = '20'+raw[2].toString();
    nmi = mma.indexOf(nmr);
    nmm = (nmi.toString()).padStart(2,"0")
    let ndt = new Date(nyy+'-'+nmm+'-'+ndd);
    console.log(str, ndt.toISOString())
    return ndt;
};

async function tesLoop(raw){
    console.log(raw)
    let nData = await raw.data;
    console.log('all data:')
    console.log(nData);
    console.log('-------');
    await cleanData(nData);
};

async function cleanData(data){
    let cData = [];
    for (let i = 0; i < data.length; i++) {
        let kStkr = data[i]['Nomor Invoice'].slice(7,13);

        let nSPM = data[i]['Nomor Invoice'].slice(0,5);
        let tSPM = data[i]['Tanggal Invoice'];
        let bSPM = data[i]['Tanggal Invoice'].slice(3,5);

        let nSP2D = data[i]['Nomor SP2D'];
        let tSP2D = data[i]['Tanggal SP2D'];
        let bSP2D = data[i]['Tanggal SP2D'].slice(3,5);
        let vSP2D = data[i]['Nilai SP2D Ekuivalen'];

        let jSPM = data[i]['Jenis SPM'];
        let uSPM = data[i]['Deskripsi'].toLowerCase();

        if(uSPM.includes("belanja pegawai")){
            aSPM = '51';
         } else if(uSPM.includes("belanja barang")){
            aSPM = '52';
         } else if(uSPM.includes("belanja modal")){
            aSPM = '53';
         } else {
            aSPM = '99 - Lainnya';
         };



        cData.push({kStkr, nSPM, tSPM, bSPM, jSPM, aSPM, nSP2D, tSP2D, bSP2D, vSP2D}); 
        
        let nRow = document.createElement('tr');
        nRow.setAttribute('id', nSPM);
        nRow.innerHTML =`
        <td>'${kStkr}</td>
        <td>'${nSPM}</td>
        <td>${tSPM}</td>
        <td>${bSPM}</td>
        <td>${nSP2D}</td>
        <td>${tSP2D}</td>
        <td>${jSPM}</td>
        <td>${aSPM}</td>
        <td>${vSP2D}</td>
        `
        document.getElementById('cleanData').appendChild(nRow);
    };
    localStorage.setItem('rawData', JSON.stringify(cData));
};


async function filJenis(data, par){
    let resJenis = [];
    for (let i = 0; i < data.length; i++) {
        if(par === 'SEMUA'){
            resJenis.push(data[i]);
        } else {
            if(data[i].jSPM === par){
                resJenis.push(data[i]);
            };
        };
    };
    return resJenis;
};
async function filBulan(data, par){
    let resBulan = [];
    for (let i = 0; i < data.length; i++) {
        if(par === 'SEMUA'){
            resBulan.push(data[i]);
        } else {
            if(data[i].bSPM === par){
                resBulan.push(data[i]);
            };
        };
    };
    return resBulan;
};
async function filAkun(data, par){
    let resAkun = [];
    for (let i = 0; i < data.length; i++) {
        if(par === 'SEMUA'){
            resAkun.push(data[i]);
        } else {
            console.log(data[i].aSPM);
            console.log(par)
            if(data[i].aSPM === par){
                resAkun.push(data[i]);
            };
        };
    };
    return resAkun;
};

async function doFilter(){
    let rData = localStorage.getItem('rawData');
    let cData = JSON.parse(rData);

    let parJenis = document.getElementById('jenisSPM').value;
    let parBulan = document.getElementById('bulanSPM').value;
    let parAkun = document.getElementById('akunSPM').value;

    let resJenis = await filJenis(cData, parJenis);
    let resBulan = await filBulan(resJenis, parBulan);
    let resFinal = await filAkun(resBulan, parAkun);

    console.log(resJenis);
    console.log(resBulan);
    console.log(resFinal);

    document.getElementById('cleanData').innerHTML = '';
    for (let i = 0; i < resFinal.length; i++) {
        let nRow = document.createElement('tr');
        nRow.setAttribute('id', resFinal[i].nSPM);
        nRow.innerHTML =`
        <td>'${resFinal[i].kStkr}</td>
        <td>'${resFinal[i].nSPM}</td>
        <td>${resFinal[i].tSPM}</td>
        <td>${resFinal[i].bSPM}</td>
        <td>${resFinal[i].nSP2D}</td>
        <td>${resFinal[i].tSP2D}</td>
        <td>${resFinal[i].jSPM}</td>
        <td>${resFinal[i].aSPM}</td>
        <td>${resFinal[i].vSP2D}</td>
        `
        document.getElementById('cleanData').appendChild(nRow);
    };
    localStorage.setItem('filData', JSON.stringify(resFinal));
};


async function prepFilter(){
    var jSPM = document.getElementById("jenisSPM");
    var bSPM = document.getElementById("bulanSPM");
    var aSPM = document.getElementById("akunSPM");

    for (let i = 0; i < jenisSPM.length; i++) {
        var nOps = document.createElement("option");
        nOps.value = jenisSPM[i];
        nOps.innerHTML = jenisSPM[i];
        jSPM.add(nOps);
    };
    for (let j = 0; j < bulanSPM.length; j++) {
        var nOps = document.createElement("option");
        nOps.value = bulanSPM[j];
        nOps.innerHTML = bulanSPM[j];
        bSPM.add(nOps);
    };
    for (let k = 0; k < akunSPM.length; k++) {
        var nOps = document.createElement("option");
        nOps.value = akunSPM[k];
        nOps.innerHTML = akunSPM[k];
        aSPM.add(nOps);
    };
};
async function showFilter(){
    document.getElementById('filterBox').classList.toggle('hide');
};


function ExportToExcel(type, fn, dl) {
    var elt = document.getElementById('cleanTable');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || ('filterSP2D.' + (type || 'xlsx')));
 };

 prepFilter();