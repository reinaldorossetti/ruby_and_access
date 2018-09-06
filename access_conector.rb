require 'win32ole'

mdb_file = __dir__ + '\\db\\exemplo.accdb'

class AccessDb
  attr_accessor :mdb, :connection, :data, :fields, :catalog

  def initialize(mdb=nil)
    @mdb = mdb
    @connection = nil
    @data = nil
    @fields = nil
  end

  def open
    connection_string =  'Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;Data Source='
    connection_string << @mdb
    WIN32OLE.codepage = WIN32OLE::CP_UTF8
    @connection = WIN32OLE.new('ADODB.Connection')
    @connection.Open(connection_string)
    @catalog = WIN32OLE.new("ADOX.Catalog")
    @catalog.ActiveConnection = @connection
  end

  def query(sql)
    recordset = WIN32OLE.new('ADODB.Recordset')
    # puts recordset.ole_methods # Mostra todos os Metodos.
    recordset.Open(sql, @connection)
    @fields = []
    recordset.Fields.each do |field|
      @fields << field.Name
    end
    begin
      @data = recordset.GetRows.transpose
    rescue
      @data = []
    end
    #recordset.Close(0)
    [@fields, @data]
  end

  def tables
    tables = []
    @catalog.tables.each {|t| tables << t.name if t.type == "TABLE" }
    tables
  end

  def execute(sql)
    @connection.Execute(sql)
  end

  def quit
    @connection.Close
  end
end

db = AccessDb.new(mdb_file)
db.open
puts db.query("Select * from relatorio")
puts db.tables
db.quit # Fecha a sessão aberta



