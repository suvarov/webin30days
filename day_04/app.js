function getData() {
  var conn = new ActiveXObject("ADODB.Connection") ;

  var connStr = "Data Source=<server>;Initial Catalog=CityDelivery;User ID=<username>;Password=<password>;Provider=SQLOLEDB";

  conn.Open(connStr);
  var rs = new ActiveXObject("ADODB.Recordset");

  rs.Open("SELECT * FROM <table>", conn);
  rs.MoveFirst
  while(!rs.eof)
  {
     document.write('<span>' + rs.fields(1) + '</span><br>');
     rs.movenext;
  }

  rs.close;
  conn.close;
}
