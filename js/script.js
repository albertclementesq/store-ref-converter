var tempData; //Temporary array to store each row of excel
var fileUpload = document.getElementById('fileUpload'); //file input selector

//Upload Excel file
function Upload() {
    //validate if file is a valid Excel file
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;

    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof(FileReader) != 'undefined') {

            var reader = new FileReader();

            //For browsers that are not IE
            if (reader.readAsBinaryString) {
                reader.onload = function(e) {
                    ProcessExcel(e.target.result);
                };
                reader.readAsBinaryString(fileUpload.files[0]);
            } else {
                //Only IE
                reader.onload = function(e) {
                    var data = "";
                    var bytes = new Uint8Array(e.target.result);
                    for (var index = 0; index < bytes.byteLength; index++) {
                        data += String.fromCharCode(bytes[index]);
                    }
                    ProcessExcel(data);
                };
                reader.readAsArrayBuffer(fileUpload.files[0]);
            }
        } else {
            UIkit.notification({
                message: 'This browser does not support HTML5.',
                status: 'danger',
                pos: 'top-center',
                timeout: 5000
            });
        }
    } else {
        UIkit.notification({
            message: 'Please, select a file first!',
            status: 'warning',
            pos: 'top-center',
            timeout: 5000
        });
    }
};

//Process Excel uploaded file and create a table to show info on it.
function ProcessExcel(data) {
    //Read Excel file data
    var workbook = XLSX.read(data, {type: 'binary'});

    //Fetch the name of the first sheet
    var firstSheet = workbook.SheetNames[0];

    //Read all rows from first sheet into a JSON array
    var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);
    
    //Create table
    var table = document.createElement('table');
    table.classList.add('uk-table', 'uk-table-striped', 'uk-table-responsive');
    
    //Add a header row
    var thead = document.createElement('thead');

    //Add header cells
    var row = thead.insertRow(-1);
    var headerCell = document.createElement('th');
    headerCell.innerHTML = 'Store Ref';
    row.appendChild(headerCell);

    headerCell = document.createElement('th');
    headerCell.innerHTML = 'Store Name';
    row.appendChild(headerCell);

    headerCell = document.createElement('th');
    headerCell.innerHTML = 'Store Telephone';
    row.appendChild(headerCell);

    headerCell = document.createElement('th');
    headerCell.innerHTML = 'Store Email';
    row.appendChild(headerCell);

    table.appendChild(thead);

    //Add each store to a row of the table
    var tbody = document.createElement('tbody');
    //Add data from excel file to the table
    for (var index = 0; index < excelRows.length; index++) {
        //tempData.push(excelRows[index]);
        //Add data row
        var row = tbody.insertRow(-1);

        //Add cells
        var cell = row.insertCell(-1);
        cell.innerHTML = excelRows[index].store_ref;
        
        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[index].store_name;
        
        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[index].store_phone;
        
        cell = row.insertCell(-1);
        cell.innerHTML = excelRows[index].store_email;
    }
    table.appendChild(tbody);

    //Create table and append to container
    var dvExcel = document.getElementById('dvExcel');
    dvExcel.innerHTML = '';
    dvExcel.appendChild(table);
    tempData = excelRows;
}

(function () {
    var textFile = null;

    var makeTxtFile = function(storesContact) {
        var stores = []; //Array to store entries after that create from it txt file

        function cleanEmail(email) {
            var cleaningString = /[^a-zA-Z0-9-_.]/g;
            var at = email.split("@");
            var username = at[0].replace(cleaningString, '');
            var domain = at[1].replace(cleaningString, '');

            return cleanedEmail = username+"@"+domain;
        }
        //clean each value of every object
        storesContact.forEach(entry => {
            var storeRef = entry.store_ref.toString().toLowerCase().replace(/[^a-zA-Z0-9_-\s]/g, '').replace(/(\s)/g, '_').replace(/(-)/g, '_');
            var storeName = entry.store_name.toString().replace(/[^a-zA-Z0-9-\s]/g, '');
            var storePhone = entry.store_phone.toString().replace(/([^0-9])/g, '').replace(/(^34)/g, '').replace(/(^0034)/g, '');
            var storeEmail = cleanEmail(entry.store_email.toLowerCase());
            
            //create final object for Ruby as a simple string
            var storeCreate = `Store.create!(\r\tmerchant: merchant,\r\treference: "${storeRef}",\r\tdata: {"name"=>"${storeName}", "store_phone"=>"${storePhone}", "store_email"=>"${storeEmail}"},\r\town: true)`;
            
            //Add string as element to array
            stores.push(storeCreate);
        });

        //Join each store element of the array
        var storesToString = stores.join('\n\n');
        
        //Show store/s into textarea
        var txtArea = document.getElementById("storesTxtArea");
        txtArea.value = storesToString;

        //Create text file
        var data = new Blob([storesToString], {type: 'text/plain'});
                
        if (textFile !== null) {
            window.URL.revokeObjectURL(textFile);
        }

        textFile = window.URL.createObjectURL(data);
        return textFile;
    }

    //Listen to user action to create and download file
    var convertData = document.getElementById('convertData');

    convertData.addEventListener('click', function() {
        var downloadLink = document.getElementById('downloadLink');
        downloadLink.href = makeTxtFile(tempData);
        downloadLink.style.display = 'block';
    }, false);
})();