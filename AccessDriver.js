
/* Declare Database Setup Options Here
************************************************/
var dbOptions = {
  dbPath: "C:\\Users\\alverer\\Downloads\\ngReact-master\\VBA-JS\\sample.accdb",
  dbUserID: "",
  dbPassword: ""
};

/* MS Access API
 ************************************************/
var MSAccess = function(dbOptions) {
  this.dbOptions = dbOptions;
  this.myConn = new ActiveXObject("ADODB.Connection");
  this.connStr = "";
  this.sessionStr = "";
  this.connOption;
  var providers = ['Microsoft.ACE.OLEDB.12.0'],
    connError = [];
  // Test for connectivity
  for (var i = 0, x = providers.length; i <= x; i++) {
    var testConn = new ActiveXObject("ADODB.Connection");
    if (this.dbOptions.dbPassword.length > 1) {
      this.connStr = "Provider=" + providers[i] + ";Data Source=" + this.dbOptions.dbPath + ";Jet OLEDB:Database Password=" + this.dbOptions.dbPassword + ";";
    } else {
      this.connStr = "Provider=" + providers[i] + ";Data Source=" + this.dbOptions.dbPath + ";";
    }
    try {
      testConn.Open(this.connStr);
      if (testConn.State === 1) {
        this.connOption = i;
        this.sessionStr = this.connStr;
      }
    } catch (error) {
      testConn = undefined;
      connError.push(providers[i]);
    }
  }
  alert("Connection Successful with: " + providers[this.connOption]);
  // Start Connections
  if (this.connOption !== false) {
    this.myConn.Open(this.sessionStr);
    alert("Connection Successful");
  } else {
    this.myConn = undefined;
    alert("Connection Test Failed All Providers");
  }
};
/* MS Access API Prototypes
 ************************************************/
MSAccess.prototype = {
  state: function() {
    var status = "";
    switch (this.myConn.State) {
      case 0:
        status = "Not Connected";
        break;
      case 1:
        status = "Connected to " + this.dbOptions.dbPath;
        break;
      default:
    }
    return status;
  },
  runQuery: function(sql) {
    var results = [],
      fieldNames = [],
      rs = new ActiveXObject("ADODB.Recordset");
    // Send SQL to Build Recordset
    rs.Open(sql, this.myConn);
    // Collect FieldNames
    for (var i = 0, x = rs.Fields.Count; i < x; i++) {
      fieldNames.push(rs.Fields(i).name);
    }
    // Build Data Collection
    while (rs.eof === false) {
      var record = {};
      for (var z = 0, y = fieldNames.length; z < y; z++) {
        record[fieldNames[z]] = String(rs.Fields(z));
      }
      results.push(record);
      rs.MoveNext;
    }
    rs.Close();
    return results;
  },
  alertResults: function(sql) {
    var res = this.runQuery(sql),
      resultStr = "";
    for (var i = 0, x = res.length; i < x; i++) {
      // Print fieldnames if at beginning
      if (i === 0) {
        for (var r in res[i]) {
          resultStr += r + "\t";
        }
        // Move to new line
        resultStr += "\n";
        for (var p in res[i]) {
          resultStr += res[i][p] + "\t";
        }
        // Move to new line
        resultStr += "\n";
      } else {
        for (var q in res[i]) {
          resultStr += res[i][q] + "\t";
        }
        // Move to new line
        resultStr += "\n";
      }
    }
    alert(resultStr);
  },
  tblStruct: function(tbl) {}
};

var myApp = new MSAccess(dbOptions);

//myApp.state
//myApp.runQuery("insert into studs ([studName]) values ('sample from js')")
myApp.alertResults("select * from studs")
