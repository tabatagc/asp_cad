<!--#include file= "../net/conn.net"-->
<!--#include file= "../net/util_per.net"-->
<!--#include file= "../net/utilitarios.net"-->
<%
nOpcaoCad_Alt_Exc =  Request.QueryString("Opcao")
if nOpcaoCad_Alt_Exc =  1 then 'Casdatro, Alteração, Exclusão de Tipo de Cliente
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
	if nOpcao =  1 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI02" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Tipo Cliente !!!")
		  end if
  		  cTpc_Var_Descricao =  Replace( Request.Form("Tpc_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_tipo_cliente '0','" & cTpc_Var_Descricao & "','" & dData & "','" & dData & "','1'" 
		  Set Sp_Tipo_Cliente =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Tipo_Cliente =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Adicionando Tipo de Cliente" )
	elseif nOpcao =  2 then 
		  if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI02" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Tipo Cliente !!!")
		  end if
		  nId_Tipo_Cliente =  Request.QueryString("Id_Tipo_Cliente") 
  		  cTpc_Var_Descricao =  Replace( Request.Form("Tpc_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_tipo_cliente  '" & nId_Tipo_Cliente & "','" & cTpc_Var_Descricao & "','" & dData & "','" & dData & "','2'"
		  Set Sp_Tipo_Cliente =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Tipo_Cliente =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Alterando Tipo de Cliente Nº " & nId_Tipo_Cliente )
	elseif nOpcao =  3 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI02" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Tipo Cliente !!!")
		  end if
		  nId_Tipo_Cliente =  Request.QueryString("Id_Tipo_Cliente")
		  cSql =  "sp_exc_Tipo_Cliente '" & nId_Tipo_Cliente & "'"
		  Set Sp_Tipo_Cliente =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Tipo_Cliente =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Excluindo Tipo de Cliente Nº " & nId_Tipo_Cliente )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_tipo_Cliente.asp")	
elseif nOpcaoCad_Alt_Exc =  2 then 'Casdatro, Alteração, Exclusão de Segmento de Mercado
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI03" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Segmento de Mercado !!!")
		  end if
  		  cSgc_Var_Descricao =  Replace( Request.Form("Sgc_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_segmento_mercado  '0','" & cSgc_Var_Descricao & "','" & dData & "','" & dData & "','1'" 
		  Set Sp_Segmento_Mercado =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Segmento_Mercado =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Adicionando Segmento de Mercado" )
	elseif nOpcao =  2 then 
		  if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI03" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Segmento de Mercado !!!")
		  end if
		  nId_Segmento_Mercado =  Request.QueryString("Id_Segmento_Mercado") 
  		  cSgc_Var_Descricao =  Replace( Request.Form("Sgc_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_segmento_mercado  '" & nId_Segmento_Mercado & "','" & cSgc_Var_Descricao & "','" & dData & "','" & dData & "','2'"
		  Set Sp_Segmento_Mercado =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Segmento_Mercado =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Alterando Segmento de Mercado Nº " & nId_Segmento_Mercado )
	 elseif nOpcao =  3 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI03" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Segmento de Mercado !!!")
		  end if
		  nId_Segmento_Mercado =  Request.QueryString("Id_Segmento_Mercado")
		  cSql =  "sp_exc_Segmento_Mercado '" & nId_Segmento_Mercado & "'"
		  Set Sp_Segmento_Mercado =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Segmento_Mercado =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Excluindo Segmento de Mercado Nº " & nId_Segmento_Mercado )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_Segmento_Mercado.asp")
elseif nOpcaoCad_Alt_Exc =  4 then 'Casdatro, Alteração, Exclusão de Cliente
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		
		cCli_Var_Razao_Social =  Replace( ucase(Request.Form("Cli_Var_Razao_Social")), "'", "")
		cCli_Char_Tipo_Pessoa =  Replace( Request.Querystring("Cli_Char_Tipo_Pessoa"), "'", "")
		nSts_Int_Id_Status =  Replace( Request.Form("Sts_Int_Id_Status"), "'", "")
		nClass_Int_Id_Classificacao =  Replace( Request.Form("Classificacao"), "'", "")
		cCli_Var_Fantasia =  Replace( ucase(Request.Form("Cli_Var_Fantasia")), "'", "")
		nVen_Int_Id_Vendedor =  Replace( Request.Form("Ven_Int_Id_Vendedor"), "'", "")
		nSgc_Int_Id_Segmento_Mercado =  Replace( Request.Form("Sgc_Int_Id_Segmento_Mercado"), "'", "")
		nTpc_Int_Id_Tipo_Cliente =  Replace( Request.Form("Tpc_Int_Id_Tipo_Cliente"), "'", "")
		cCli_Var_CNPJ = Replace( Replace( Replace( Replace( Replace( Request.Form("Cli_Var_CNPJ"), "'", "" ), ",", "" ), "/" , "" ), "." , "" ), "-" , "" )

		cSql = "select count (*) as soma from t_cliente where cli_var_cnpj = '" & cCli_Var_CNPJ & "'"
		set rs = cConexao.Execute (cSql)
			if rs("soma") > 0 then
			response.Redirect("../net/msg.asp?Msg=CNPJ/CPF já cadastrado !!!")
			end if

		cCli_Var_IE =  Replace( Request.Form("Cli_Var_IE"), "'", "")
		cCli_Var_CCM =  Replace( Request.Form("Cli_Var_CCM"), "'", "")
		nCli_Dec_Limite_Credito =  Replace( Replace( Request.Form("Cli_Dec_Limite_Credito"), "'", "" ), ",", "." )
		if ( nCli_Dec_Limite_Credito = Empty ) then
			nCli_Dec_Limite_Credito = 0
		end if
		cCli_Char_DDD =  Replace( Request.Form("Cli_Char_DDD"), "'", "")
		cCli_Var_Telefone =  Replace( Request.Form("Cli_Var_Telefone"), "'", "")
		cCli_Var_Ramal =  Replace( Request.Form("Cli_Var_Ramal"), "'", "")
		cCli_Var_Fax =  Replace( Request.Form("Cli_Var_Fax"), "'", "")
		cCli_Var_Ramal_Fax =  Replace( Request.Form("Cli_Var_Ramal_Fax"), "'", "")
		cCli_Var_Home_Page =  Replace( Request.Form("Cli_Var_Home_Page"), "'", "")		
		cCli_Text_Obs =  Replace( Request.Form("Cli_Text_Obs"), "'", "")				
		cCli_Text_Obs_Cadastro =  Replace( Request.Form("Cli_Text_Obs_Cadastro"), "'", "")
		nPa_Int_Id_Pais = request.Form("Pa_Int_Id_Pais")

		if ( request.QueryString("Tipo") <> "forn" and cint(nPa_Int_Id_Pais) = 1 ) then
			if VerificaCnpjCpf( cCli_Var_CNPJ, 0 ) then
				if cCli_Char_Tipo_Pessoa = "F" then
					cMsg = "CPF já cadastrado !!!"
				else 
					cMsg = "CNPJ já cadastrado !!!"
				end if
				Response.Redirect("../net/msg.asp?Msg="&cMsg)
			end if
		end if
		
		cCli_char_flag_sm = Request.Form("Cli_char_flag_sm")
		cCli_char_flag_representante = Request.Form("Cli_char_flag_representante")
		
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		
		cSql =  "sp_cad_alt_cliente  '0','" & nVen_Int_Id_Vendedor & "','" & nTpc_Int_Id_Tipo_Cliente & "','" & nSgc_Int_Id_Segmento_Mercado & "','" & nSts_Int_Id_Status 
		cSql =  cSql & "','" & cCli_Var_Razao_Social & "','" & cCli_Var_Fantasia & "','" & cCli_Char_Tipo_Pessoa & "','" & cCli_Var_CNPJ & "','" & cCli_Var_IE & "','" & cCli_Var_CCM 
		cSql =  cSql & "','" & nCli_Dec_Limite_Credito & "','" & cCli_Char_DDD & "','" & cCli_Var_Telefone & "','" & cCli_Var_Ramal & "','" & cCli_Var_Fax & "','" & cCli_Var_Ramal_Fax 
		cSql =  cSql & "','" & cCli_Var_Home_Page & "','" & cCli_Text_Obs & "','" & cCli_Text_Obs_Cadastro & "','01','"  
		cSql =  cSql & cCli_char_flag_sm & "', '" & cCli_char_flag_representante & "', '" & nClass_Int_Id_Classificacao & "' , '" & dData & "','" & dData &  "','"&nPa_Int_Id_Pais&"','1'" 

		Set Sp_Cliente =  cConexao.Execute( cSql )

		'Call TrataErroSistema("")
		Set Sp_Cliente =  Nothing
		cSql = "select @@Identity as Numero_atual"
		Set RsSql = cConexao.Execute( cSql )
		if ( RsSql.RecordCount > 0 ) and ( not RsSql.Eof ) then
			Session("Cad_Id_Cliente") = RsSql("Numero_Atual")
		end if
		RsSql.Close : Set RsSql = Nothing
' não lê a sessão acima!

		if request.QueryString("Tipo") = "forn" or (nPa_Int_Id_Pais <> 1 and cCli_Char_Tipo_Pessoa = "F") then
		ySql = "update t_cliente set cli_var_cnpj = '" & Session("Cad_Id_Cliente") & "' where cli_int_id_cliente = " & Session("Cad_Id_Cliente")
		Set RsySql = cConexao.Execute( ySql )
		RsySql.Close : Set RsySql = Nothing
		end if
		Verifica_Integridade_Cliente( Session("Cad_Id_Cliente") )

	    Call Log_Seguranca( "Modulo Cliente", "Adicionando Cliente Nº " & Session("Cad_Id_Cliente") )
'inicia o cadastro único, ENDEREÇO		
		'nCli_Int_Id_End_Cliente =  Session("Cad_Id_Cliente")
		cEdc_Var_Endereco =  Replace( Request.Form("Edc_Var_Endereco"), "'", "")
		nEdc_Int_Numero =  Replace( Request.Form("Edc_Int_Numero"), "'", "")
		cEdc_Var_Complemento =  Replace( Request.Form("Edc_Var_Complemento"), "'", "")
		cEdc_Var_Bairro =  Replace( Request.Form("Edc_Var_Bairro"), "'", "")
		cEdc_Var_Cep =  Replace( Request.Form("Edc_Var_Cep"), "'", "")
		cEdc_Var_Cidade =  Replace( Request.Form("Edc_Var_Cidade"), "'", "")
		cEdc_Char_Estado =  Replace( Request.Form("Edc_Char_Estado"), "'", "")
		cEdc_Char_Tipo_Endereco =  "F"
		
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		
		cSql =  "sp_cad_alt_end_cliente  '" & Session("Cad_Id_Cliente") & "','0','" & cEdc_Var_Endereco & "','" & nEdc_Int_Numero 
		cSql =  cSql & "','" & cEdc_Var_Complemento & "','" & cEdc_Var_Bairro & "','" & cEdc_Var_Cep & "','" & cEdc_Var_Cidade 
		cSql =  cSql & "','" & cEdc_Char_Estado & "','" & cEdc_Char_Tipo_Endereco & "','" & dData & "','" & dData &  "','1'" 
		Set Sp_End_Cliente =  cConexao.Execute( cSql )
		'Call TrataErroSistema("")
		Set Sp_End_Cliente =  Nothing
		
		cSql = "update t_cliente set cli_int_data_sistema = '" & dData & "' where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "'"
		Set RsSql = cConexao.Execute( cSql )
		'response.Write(cSql)
		'response.End()

		Set RsSql = Nothing
		
	    Call Log_Seguranca( "Modulo Cliente", "Adicionando Endereço ao Cliente nº"  & Session("Cad_Id_Cliente") & "'" )
'inicia o cadastro único CONTATO
		'nCli_Int_Id_Cliente = request.QueryString("id_cliente")
		nCtc_Int_Id_Contato = Replace( Request.Form("Ctc_Int_Id_Contato"), "'", "")
		nVen_Int_Id_Vendedor2 = Replace( Request.Form("Ven_Int_Id_Vendedor2"), "'", "")
		nSgc_Int_Id_Status = Replace( Request.Form("Sgc_Int_Id_Status"), "'", "")
		cCtc_Var_Nome = Replace( Request.Form("Ctc_Var_Nome"), "'", "")
		cCtc_Var_Cargo = Replace( Request.Form("Ctc_Var_Cargo"), "'", "")
		cCtc_Char_Sexo = Replace( Request.Form("Ctc_Char_Sexo"), "'", "")
		dCtc_Int_Data_Nascimento = Replace( Request.Form("Ctc_Int_Data_Nascimento"), "'", "")
		if dCtc_Int_Data_Nascimento <> Empty then
			xData = dCtc_Int_Data_Nascimento
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Nascimento Invalida !!!")
			end if
			dCtc_Int_Data_Nascimento = Year( xData ) & Right("0" & Month( xData ), 2 ) & Right("0" & Day( xData ), 2 )
		end if
		cCtc_Char_Ddd = Replace( Request.Form("Ctc_Char_Ddd"), "'", "")
		cCtc_Var_Telefone = Replace( Request.Form("Ctc_Var_Telefone"), "'", "")
		cCtc_Var_Ramal = Replace( Request.Form("Ctc_Var_Ramal"), "'", "")
		cCtc_Var_Fax = Replace( Request.Form("Ctc_Var_Fax"), "'", "")
		cCtc_Var_Ramal_Fax = Replace( Request.Form("Ctc_Var_Ramal_Fax"), "'", "")
		cCtc_Char_Ddd_Celular = Replace( Request.Form("Ctc_Char_Ddd_Celular"), "'", "")
		cCtc_Var_Celular = Replace( Request.Form("Ctc_Var_Celular"), "'", "")
		cCtc_Var_Email = Replace( Request.Form("Ctc_Var_Email"), "'", "")
		
		'sql = "Select count(ctc_var_email) as soma from t_contato where ctc_var_email ='" & cCtc_Var_Email & "'"
    'set rs = cconexao.execute(sql)
      '  if Cint(rs("soma")) > 0 then
		'		response.Redirect("../net/msg.asp?Msg=E-Mail já cadastrado !!!")
		'	set xrs = cconexao.execute(csql)
      '  end if
	'rs.close : set rs = nothing

		'cCtc_Var_Home_Page = Replace( Request.Form("Ctc_Var_Home_Page"), "'", "")
		cCtc_Var_Home_Page = cCli_Var_Home_Page
		cCtc_Var_Senha_Internet = Replace( Request.Form("Ctc_Var_Senha_Internet"), "'", "")
		  if ( Request.Form("Ctc_Char_Flag_Internet") = "T" ) then
			  lCtc_Char_Flag_Internet = "T"
		  else
			  lCtc_Char_Flag_Internet = "F"
		  end if
		  if ( Request.Form("Ctc_Char_Flag_PN") = "T" ) then
			  cCtc_Char_Flag_PN = "T"
		  else
			  cCtc_Char_Flag_PN = "F"
		  end if

		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		cSql =  "sp_cad_alt_contato '0','" & Session("Cad_Id_Cliente") & "','" & nVen_Int_Id_Vendedor2
		cSql =  cSql & "','" & nSgc_Int_Id_Status & "','" & cCtc_Var_Nome & "','" & cCtc_Var_Cargo & "','" & cCtc_Char_Sexo 
		cSql =  cSql & "','" & dCtc_Int_Data_Nascimento & "','" & cCtc_Char_Ddd & "','" & cCtc_Var_Telefone
		cSql =  cSql & "','" & cCtc_Var_Ramal & "','" & cCtc_Var_Fax & "','" & cCtc_Var_Ramal_Fax & "','" & cCtc_Char_Ddd_Celular & "','" & cCtc_Var_Celular 
		cSql =  cSql & "','" & cCtc_Var_Email & "','" & cCtc_Var_Home_Page & "','" & cCtc_Var_Senha_Internet 
		cSql =  cSql & "','" & lCtc_Char_Flag_Internet & "','" & dData & "','" & dData & "','" & cCtc_Char_Flag_PN & "','1'" 
		if request.Form("id_contato") = Empty then
		  Set Sp_Contato =  cConexao.Execute( cSql )
		  'Call TrataErroSistema("")
		  Set Sp_Contato =  Nothing
		Else
		cconexao.execute("insert into t_cliente_contato values ("&Session("Cad_Id_Cliente")&", "&request.form("id_contato")&")")
		end if
		cSql = "update t_cliente set cli_int_data_sistema = '" & dData & "' where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "'"
		Set RsSql = cConexao.Execute( cSql )
		Set RsSql = Nothing
		
		idctcSql = "select ctc_int_id_contato from t_contato where cli_int_id_cliente ='" & Session("Cad_Id_Cliente") & "' and ctc_var_nome = '" & cCtc_Var_Nome & "'"
		Set Rsidctc = cConexao.Execute( idctcSql  )
		contato = Rsidctc("ctc_int_id_contato")
		Set Rsidctc = Nothing
		
	    Call Log_Seguranca( "Modulo Cliente", "Adicionando Contato Nº " & contato)

		Call FechaConexaoSistema()
		
	    response.Redirect("cad_alt_cliente.asp?Opcao=2&Id_Cliente="&Session("Cad_Id_Cliente"))
	elseif nOpcao =  2 then 
		if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nCli_Int_Id_Cliente =  Session("Cad_Id_Cliente")
		cCli_Var_Razao_Social =  Replace( Request.Form("Cli_Var_Razao_Social"), "'", "")
		cCli_Char_Tipo_Pessoa =  Replace( Request.Form("Cli_Char_Tipo_Pessoa"), "'", "")
		if ( cCli_Char_Tipo_Pessoa = Empty ) then
			cCli_Char_Tipo_Pessoa = Session("aDados")(9,1)
		end if
			nCli_Int_Id_Cliente = request.form("Cli_id")
		nSts_Int_Id_Status =  Replace( Request.Form("Sts_Int_Id_Status"), "'", "")
		nClass_Int_Id_Classificacao =  Replace( Request.Form("Classificacao"), "'", "")
		cCli_Var_Fantasia =  Replace( Request.Form("Cli_Var_Fantasia"), "'", "")
		nVen_Int_Id_Vendedor =  Replace( Request.Form("Ven_Int_Id_Vendedor"), "'", "")
		if nSts_Int_Id_Status = 6 then
			nVen_Int_Id_Vendedor = 33
			
		Set Cryptor = Server.CreateObject("AspCrypt.Crypt")
		   cStrSalt = "0123456789"
		   cStrValue = "kuerfhkaeufyxuj"
		   cSenha = Cryptor.Crypt( cStrSalt, cStrValue)
		Set Cryptor = Nothing
			
			ctcsql = "update t_contato set ctc_char_flag_internet = 'F', ctc_var_senha_internet = '" & cSenha & "' where cli_int_id_cliente  = '" & nCli_Int_Id_Cliente & "'"
			cConexao.Execute( ctcsql )
		end if
		if ( Replace( Request.Form("Ven_Int_Id_Vendedor"), "'", "") = Empty ) then
		   nVen_Int_Id_Vendedor = Session("aDados")(11,1)
		end if
		nSgc_Int_Id_Segmento_Mercado =  Replace( Request.Form("Sgc_Int_Id_Segmento_Mercado"), "'", "")
		cCli_Var_CNPJ =  Replace( Request.Form("cCli_Var_CNPJ"), "'", "")
		nTpc_Int_Id_Tipo_Cliente =  Replace( Request.Form("Tpc_Int_Id_Tipo_Cliente"), "'", "")
		'cCli_Var_CNPJ =  Session("aDados")(1,1) 'Replace( Request.Form("Cli_Var_CNPJ"), "'", "")
		cCli_Var_IE =  Replace( Request.Form("Cli_Var_IE"), "'", "")
'		cCli_Var_IE =  Session("aDados")(2,1) 'Replace( Request.Form("Cli_Var_IE"), "'", "")
		cCli_Var_CCM =  Replace( Request.Form("Cli_Var_CCM"), "'", "")
'		nCli_Dec_Limite_Credito = Replace( Replace( Request.Form("Cli_Dec_Limite_Credito"), "'", "" ), ",", "." )
		if ( nCli_Dec_Limite_Credito = Empty ) then
			nCli_Dec_Limite_Credito = 0
		end if
		if Trim( (Replace( Request.Form("Cli_Char_DDD"), "'", "") ) = empty) then 
			cCli_Char_DDD = Session("aDados")(3,1)
		else
			cCli_Char_DDD =  Replace( Request.Form("Cli_Char_DDD"), "'", "")
		end if 
		if Trim( (Replace( Request.Form("Cli_Var_Telefone"), "'", "") ) = empty) then 
			cCli_Var_Telefone = Session("aDados")(4,1)
		else
			cCli_Var_Telefone =  Replace( Request.Form("Cli_Var_Telefone"), "'", "")
		end if 
		if Trim( (Replace( Request.Form("Cli_Var_Ramal"), "'", "") ) = empty) then 
			cCli_Var_Ramal = Session("aDados")(5,1)
		else
			cCli_Var_Ramal =  Replace( Request.Form("Cli_Var_Ramal"), "'", "")
		end if 
		if Trim( (Replace( Request.Form("Cli_Var_Fax"), "'", "") ) = empty) then 
			cCli_Var_Fax = Session("aDados")(6,1)
		else
			cCli_Var_Fax =  Replace( Request.Form("Cli_Var_Fax"), "'", "")
		end if 
		if Trim( (Replace( Request.Form("Cli_Var_Ramal_Fax"), "'", "") ) = empty) then 
			cCli_Var_Ramal_Fax = Session("aDados")(7,1)
		else
			cCli_Var_Ramal_Fax =  Replace( Request.Form("Cli_Var_Ramal_Fax"), "'", "")
		end if 
		if Trim( (Replace( Request.Form("Cli_Dec_Limite_Credito"), "'", "") ) = empty) then 
			Cli_Dec_Limite_Credito = Session("aDados")(8,1)
		else
			nCli_Dec_Limite_Credito = Replace( Replace( Request.Form("Cli_Dec_Limite_Credito"), "'", "" ), ",", "." )
		end if 
' validação de cnpj
		if Trim( (Replace( Request.Form("Cli_Char_Tipo_Pessoa"), "'", "") ) = empty) then 
			Cli_Char_Tipo_Pessoa = Session("aDados")(9,1)
		else
				if VerificaCnpjCpf( cCli_Var_CNPJ, nCli_Int_Id_Cliente ) then
					if cCli_Char_Tipo_Pessoa = "F" then
						cMsg = "CPF já cadastrado !!!"
					else 
						cMsg = "CNPJ já cadastrado !!!"
					end if
					Response.Redirect("../net/msg.asp?Msg="&cMsg)
				end if
		end if
'acaba validação de cnpj		
		cCli_Var_Home_Page =  Replace( Request.Form("Cli_Var_Home_Page"), "'", "")		
		cCli_Text_Obs =  Replace( Request.Form("Cli_Text_Obs"), "'", "")				
		cCli_Text_Obs_Cadastro =  Replace( Request.Form("Cli_Text_Obs_Cadastro"), "'", "")
		
		cCli_char_flag_sm = Request.Form("Cli_char_flag_sm")
		cCli_char_flag_representante = Request.Form("Cli_char_flag_representante")
		nPa_Int_Id_Pais = request.Form("Pa_Int_Id_Pais")
		'response.Write(nPa_Int_Id_Pais)
		'response.End()				
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		cSql =  "sp_cad_alt_cliente " & "'" & nCli_Int_Id_Cliente & "','" & nVen_Int_Id_Vendedor & "','" & nTpc_Int_Id_Tipo_Cliente & "','" & nSgc_Int_Id_Segmento_Mercado & "','" & nSts_Int_Id_Status 
		cSql =  cSql & "','" & cCli_Var_Razao_Social & "','" & cCli_Var_Fantasia & "','" & cCli_Char_Tipo_Pessoa & "','" & cCli_Var_CNPJ & "','" & cCli_Var_IE & "','" & cCli_Var_CCM 
		cSql =  cSql & "','" & nCli_Dec_Limite_Credito & "','" & cCli_Char_DDD & "','" & cCli_Var_Telefone & "','" & cCli_Var_Ramal & "','" & cCli_Var_Fax & "','" & cCli_Var_Ramal_Fax 
		cSql =  cSql & "','" & cCli_Var_Home_Page & "','" & cCli_Text_Obs & "','" & cCli_Text_Obs_Cadastro & "','01','" 
		cSql =  cSql & cCli_char_flag_sm & "', '" & cCli_char_flag_representante & "', '" & nClass_Int_Id_Classificacao & "', '" & dData & "','" & dData &  "','"&nPa_Int_Id_Pais&"','2'" 
		'response.Write(cSql)
		'response.End()
		Set Sp_Cliente =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Cliente =  Nothing
		Verifica_Integridade_Cliente( nCli_Int_Id_Cliente )
	    Call Log_Seguranca( "Modulo Cliente", "Alterando Cliente Nº " & nCli_Int_Id_Cliente )
		Call FechaConexaoSistema()
	    response.Redirect("cad_alt_cliente.asp?Opcao=2&Id_Cliente="&nCli_Int_Id_Cliente)
    elseif nOpcao =  3 then
		if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
	    nCli_Int_Id_Cliente =  Request.QueryString("Id_Cliente")
	    cSql =  "exec sp_exc_Cliente '" & nCli_Int_Id_Cliente & "'"
	    Set Sp_Cliente =  cConexao.Execute( cSql )
	    Call TrataErroSistema("")
	    Set Sp_Cliente =  Nothing
	    Call Log_Seguranca( "Modulo Cliente", "Excluindo Cliente Nº " & Session("Cad_Id_Cliente") )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_cliente.asp")
elseif nOpcaoCad_Alt_Exc =  5 then 'Casdatro, Alteração, Exclusão de Endereço Cliente
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
	if nOpcao =  1 then
		if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		'verifica se o cliente já possui endereço de faturamento pois so pode ter um
		set rs_verifica = cconexao.execute("Select cli_int_id_Cliente from t_end_cliente where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "' and edc_char_tipo_endereco = 'F'")

		nCli_Int_Id_End_Cliente =  Session("Cad_Id_Cliente")
		cEdc_Var_Endereco =  Replace( Request.Form("Edc_Var_Endereco"), "'", "")
		nEdc_Int_Numero =  Replace( Request.Form("Edc_Int_Numero"), "'", "")
		cEdc_Var_Complemento =  Replace( Request.Form("Edc_Var_Complemento"), "'", "")
		cEdc_Var_Bairro =  Replace( Request.Form("Edc_Var_Bairro"), "'", "")
		cEdc_Var_Cep =  Replace( Request.Form("Edc_Var_Cep"), "'", "")
		cEdc_Var_Cidade =  Replace( Request.Form("Edc_Var_Cidade"), "'", "")
		cEdc_Char_Estado =  Replace( Request.Form("Edc_Char_Estado"), "'", "")
		cEdc_Char_Tipo_Endereco =  Replace( Request.Form("Edc_Char_Tipo_Endereco"), "'", "")
		
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		
		if not rs_verifica.eof and cEdc_Char_Tipo_Endereco = "F" then
			response.Redirect("../net/msg.asp?Msg=Cliente já possui endereço de Faturamento !!!")
		end if
		
		cSql =  "sp_cad_alt_end_cliente  '" & nCli_Int_Id_End_Cliente & "','0','" & cEdc_Var_Endereco & "','" & nEdc_Int_Numero 
		cSql =  cSql & "','" & cEdc_Var_Complemento & "','" & cEdc_Var_Bairro & "','" & cEdc_Var_Cep & "','" & cEdc_Var_Cidade 
		cSql =  cSql & "','" & cEdc_Char_Estado & "','" & cEdc_Char_Tipo_Endereco & "','" & dData & "','" & dData &  "','1'" 
		Set Sp_End_Cliente =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_End_Cliente =  Nothing
		
		cSql = "update t_cliente set cli_int_data_sistema = '" & dData & "' where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "'"
		Set RsSql = cConexao.Execute( cSql )
		Set RsSql = Nothing
		
	    Call Log_Seguranca( "Modulo Cliente", "Adicionando Endereço ao Cliente nº"  & Session("Cad_Id_Cliente") & "'" )
	elseif nOpcao =  2 then 
		if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nEdc_Int_Id_End_Cliente =  Trim( Request.QueryString("Id_End_Cliente") )
		nCli_Int_Id_End_Cliente =  Trim( Request.QueryString("Id_Cliente") )
		cEdc_Var_Endereco =  Replace( Request.Form("Edc_Var_Endereco"), "'", "")
		nEdc_Int_Numero =  Replace( Request.Form("Edc_Int_Numero"), "'", "")
		cEdc_Var_Complemento =  Replace( Request.Form("Edc_Var_Complemento"), "'", "")
		cEdc_Var_Bairro =  Replace( Request.Form("Edc_Var_Bairro"), "'", "")
		cEdc_Var_Cep =  Replace( Request.Form("Edc_Var_Cep"), "'", "")
		cEdc_Var_Cidade =  Replace( Request.Form("Edc_Var_Cidade"), "'", "")
		cEdc_Char_Estado =  Replace( Request.Form("Edc_Char_Estado"), "'", "")
		cEdc_Char_Tipo_Endereco =  Replace( Request.Form("Edc_Char_Tipo_Endereco"), "'", "")
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )		
		cSql =  "sp_cad_alt_end_cliente  '" & nCli_Int_Id_End_Cliente & "','" & nEdc_Int_Id_End_Cliente & "','" & cEdc_Var_Endereco & "','" & nEdc_Int_Numero 
		cSql =  cSql & "','" & cEdc_Var_Complemento & "','" & cEdc_Var_Bairro & "','" & cEdc_Var_Cep & "','" & cEdc_Var_Cidade 
		cSql =  cSql & "','" & cEdc_Char_Estado & "','" & cEdc_Char_Tipo_Endereco & "','" & dData & "','" & dData &  "','2'" 
		  Set Sp_End_Cliente =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_End_Cliente =  Nothing
		cSql = "update t_cliente set cli_int_data_sistema = '" & dData & "' where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "'"
		Set RsSql = cConexao.Execute( cSql )
		Set RsSql = Nothing

	    Call Log_Seguranca( "Modulo Cliente", "Alterando Endereço Cliente Nº " & nEdc_Int_Id_End_Cliente )
    elseif nOpcao =  3 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI01" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		  end if
		  nEdc_Int_Id_End_Cliente =  Request.QueryString("Id_End_Cliente")
		  cSql =  "sp_exc_End_Cliente '" & nEdc_Int_Id_End_Cliente & "'"
		  Set Sp_End_Cliente =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_End_Cliente =  Nothing
	    Call Log_Seguranca( "Modulo Cliente", "Excluindo Endereço Cliente Nº " & nEdc_Int_Id_End_Cliente )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_end_cliente.asp?Id_Cliente="&Session("Cad_id_Cliente"))
elseif nOpcaoCad_Alt_Exc =  6 then 'Casdatro, Alteração, Exclusão de Contato
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
	if nOpcao =  1 then
		if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nCli_Int_Id_Cliente = request.QueryString("id_cliente")
		nCtc_Int_Id_Contato = Replace( Request.Form("Ctc_Int_Id_Contato"), "'", "")
		nVen_Int_Id_Vendedor = Replace( Request.Form("Ven_Int_Id_Vendedor"), "'", "")
		nSgc_Int_Id_Status = Replace( Request.Form("Sgc_Int_Id_Status"), "'", "")
		cCtc_Var_Nome = Replace( Request.Form("Ctc_Var_Nome"), "'", "")
		cCtc_Var_Cargo = Replace( Request.Form("Ctc_Var_Cargo"), "'", "")
		cCtc_Char_Sexo = Replace( Request.Form("Ctc_Char_Sexo"), "'", "")
		dCtc_Int_Data_Nascimento = Replace( Request.Form("Ctc_Int_Data_Nascimento"), "'", "")
		if dCtc_Int_Data_Nascimento <> Empty then
			xData = dCtc_Int_Data_Nascimento
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Nascimento Invalida !!!")
			end if
			dCtc_Int_Data_Nascimento = Year( xData ) & Right("0" & Month( xData ), 2 ) & Right("0" & Day( xData ), 2 )
		end if
		cCtc_Char_Ddd = Replace( Request.Form("Ctc_Char_Ddd"), "'", "")
		cCtc_Var_Telefone = Replace( Request.Form("Ctc_Var_Telefone"), "'", "")
		cCtc_Var_Ramal = Replace( Request.Form("Ctc_Var_Ramal"), "'", "")
		cCtc_Var_Fax = Replace( Request.Form("Ctc_Var_Fax"), "'", "")
		cCtc_Var_Ramal_Fax = Replace( Request.Form("Ctc_Var_Ramal_Fax"), "'", "")
		cCtc_Char_Ddd_Celular = Replace( Request.Form("Ctc_Char_Ddd_Celular"), "'", "")
		cCtc_Var_Celular = Replace( Request.Form("Ctc_Var_Celular"), "'", "")
		cCtc_Var_Email = Replace( Request.Form("Ctc_Var_Email"), "'", "")
		
		sql = "Select count(ctc_var_email) as soma from t_contato where ctc_var_email ='" & cCtc_Var_Email & "'"
    set rs = cconexao.execute(sql)
        if Cint(rs("soma")) > 0 then
				response.Redirect("../net/msg.asp?Msg=E-Mail já cadastrado !!!")
			set xrs = cconexao.execute(csql)
        end if
	rs.close : set rs = nothing
	
		cCtc_Var_Home_Page = Replace( Request.Form("Ctc_Var_Home_Page"), "'", "")
		cCtc_Var_Senha_Internet = Replace( Request.Form("Ctc_Var_Senha_Internet"), "'", "")
		  if ( Request.Form("Ctc_Char_Flag_Internet") = "T" ) then
			  lCtc_Char_Flag_Internet = "T"
		  else
			  lCtc_Char_Flag_Internet = "F"
		  end if
		  if ( Request.Form("Ctc_Char_Flag_PN") = "T" ) then
			  cCtc_Char_Flag_PN = "T"
		  else
			  cCtc_Char_Flag_PN = "F"
		  end if
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		cSql =  "sp_cad_alt_contato '0','" & nCli_Int_Id_Cliente & "','" & nVen_Int_Id_Vendedor
		cSql =  cSql & "','" & nSgc_Int_Id_Status & "','" & cCtc_Var_Nome & "','" & cCtc_Var_Cargo & "','" & cCtc_Char_Sexo 
		cSql =  cSql & "','" & dCtc_Int_Data_Nascimento & "','" & cCtc_Char_Ddd & "','" & cCtc_Var_Telefone
		cSql =  cSql & "','" & cCtc_Var_Ramal & "','" & cCtc_Var_Fax & "','" & cCtc_Var_Ramal_Fax & "','" & cCtc_Char_Ddd_Celular & "','" & cCtc_Var_Celular 
		cSql =  cSql & "','" & cCtc_Var_Email & "','" & cCtc_Var_Home_Page & "','" & cCtc_Var_Senha_Internet 
		cSql =  cSql & "','" & lCtc_Char_Flag_Internet & "','" & dData & "','" & dData & "','" & cCtc_Char_Flag_PN & "','1'" 
		  Set Sp_Contato =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_Contato =  Nothing

		cSql = "update t_cliente set cli_int_data_sistema = '" & dData & "' where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "'"
		Set RsSql = cConexao.Execute( cSql )
		Set RsSql = Nothing
		
		idctcSql = "select ctc_int_id_contato from t_contato where cli_int_id_cliente ='" & Session("Cad_Id_Cliente") & "' and ctc_var_nome = '" & cCtc_Var_Nome & "'"
		Set Rsidctc = cConexao.Execute( idctcSql  )
		contato = Rsidctc("ctc_int_id_contato")
		Set Rsidctc = Nothing
		
	    Call Log_Seguranca( "Modulo Cliente", "Adicionando Contato Nº " & contato)
	elseif nOpcao =  2 then 
		if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nCli_Int_Id_Cliente = Request.QueryString("Id_Cliente")
		nCtc_Int_Id_Contato = Request.QueryString("Id_Contato")
		nVen_Int_Id_Vendedor = Replace( Request.Form("Ven_Int_Id_Vendedor"), "'", "")
		nSgc_Int_Id_Status = Replace( Request.Form("Sgc_Int_Id_Status"), "'", "")
		cCtc_Var_Nome = Replace( Request.Form("Ctc_Var_Nome"), "'", "")
		cCtc_Var_Cargo = Replace( Request.Form("Ctc_Var_Cargo"), "'", "")
		cCtc_Char_Sexo = Replace( Request.Form("Ctc_Char_Sexo"), "'", "")
		dCtc_Int_Data_Nascimento = Replace( Request.Form("Ctc_Int_Data_Nascimento"), "'", "")
		if dCtc_Int_Data_Nascimento <> Empty then
			xData = dCtc_Int_Data_Nascimento
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Nascimento Invalida !!!")
			end if
			dCtc_Int_Data_Nascimento = Year( xData ) & Right("0" & Month( xData ), 2 ) & Right("0" & Day( xData ), 2 )
		end if
		cCtc_Char_Ddd = Replace( Request.Form("Ctc_Char_Ddd"), "'", "")
		cCtc_Var_Telefone = Replace( Request.Form("Ctc_Var_Telefone"), "'", "")
		cCtc_Var_Ramal = Replace( Request.Form("Ctc_Var_Ramal"), "'", "")
		cCtc_Var_Fax = Replace( Request.Form("Ctc_Var_Fax"), "'", "")
		cCtc_Var_Ramal_Fax = Replace( Request.Form("Ctc_Var_Ramal_Fax"), "'", "")
		cCtc_Char_Ddd_Celular = Replace( Request.Form("Ctc_Char_Ddd_Celular"), "'", "")
		cCtc_Var_Celular = Replace( Request.Form("Ctc_Var_Celular"), "'", "")
		cCtc_Var_Email = Replace( Request.Form("Ctc_Var_Email"), "'", "")
		
		'sql = "Select count(ctc_var_email) as soma from t_contato where ctc_var_email ='" & cCtc_Var_Email & "' and Ctc_Int_Id_Contato <> '" & nCtc_Int_Id_Contato & "'"
    'set rs = cconexao.execute(sql)
        'if Cint(rs("soma")) > 0 then
				'dsql = "Select cli_var_cnpj, cli_var_razao_social from t_cliente as cli inner join t_contato as ctc on (cli.cli_int_id_cliente = ctc.cli_int_id_cliente) where ctc_var_email = '" & cCtc_Var_Email & "' and cli.cli_int_id_cliente <> '" & nCli_Int_Id_Cliente & "'"
				'response.Write(dsql)
				'response.End()
				'set xrs = cconexao.execute(dsql)
					'if not xrs.eof then
						'Cnpj = xrs("cli_var_cnpj")
						'Razao = xrs("cli_var_razao_social")
					'end if
				'xrs.close : set xrs = nothing
'response.Redirect("../net/msg.asp?Msg=E-Mail já cadastrado no cliente '" & Razao & "' cnpj '" & Cnpj & "' !!!")
        'end if
	'rs.close : set rs = nothing
	
		cCtc_Var_Home_Page = Replace( Request.Form("Ctc_Var_Home_Page"), "'", "")
		cCtc_Var_Senha_Internet = Replace( Request.Form("Ctc_Var_Senha_Internet"), "'", "")
		  if ( Request.Form("Ctc_Char_Flag_Internet") = "T" ) then
			  lCtc_Char_Flag_Internet = "T"
		  else
			  lCtc_Char_Flag_Internet = "F"
		  end if
		  if ( Request.Form("Ctc_Char_Flag_PN") = "T" ) then
			  cCtc_Char_Flag_PN = "T"
		  else
			  cCtc_Char_Flag_PN = "F"
		  end if
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		cSql =  "sp_cad_alt_contato '" & nCtc_Int_Id_Contato & "','" & nCli_Int_Id_Cliente & "','" & nVen_Int_Id_Vendedor
		cSql =  cSql & "','" & nSgc_Int_Id_Status & "','" & cCtc_Var_Nome & "','" & cCtc_Var_Cargo & "','" & cCtc_Char_Sexo 
		cSql =  cSql & "','" & dCtc_Int_Data_Nascimento & "','" & cCtc_Char_Ddd & "','" & cCtc_Var_Telefone
		cSql =  cSql & "','" & cCtc_Var_Ramal & "','" & cCtc_Var_Fax & "','" & cCtc_Var_Ramal_Fax & "','" & cCtc_Char_Ddd_Celular & "','" & cCtc_Var_Celular 
		cSql =  cSql & "','" & cCtc_Var_Email & "','" & cCtc_Var_Home_Page & "','" & cCtc_Var_Senha_Internet 
		cSql =  cSql & "','" & lCtc_Char_Flag_Internet & "','" & dData & "','" & dData & "','" & cCtc_Char_Flag_PN & "','2'" 
	    Set Sp_Contato =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Contato =  Nothing

		cSql = "update t_cliente set cli_int_data_sistema = '" & dData & "' where cli_int_id_cliente = '" & Session("Cad_Id_Cliente") & "'"
		Set RsSql = cConexao.Execute( cSql )
		Set RsSql = Nothing

	    Call Log_Seguranca( "Modulo Cliente", "Alterando Contato Nº " & nCtc_Int_Id_Contato )
		response.Redirect("con_contato_cliente.asp?Id_Cliente=" & nCli_Int_Id_Cliente)
		'response.Redirect("cad_alt_contato.asp?Opcao=2&Id_Contato=" & nCtc_Int_Id_Contato )
	 elseif nOpcao =  3 then
		if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI01" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nCtc_Int_Id_Contato = Request.QueryString("Id_Contato")
		aSql =  "select ctc_int_id_contato from t_agenda_cliente where ctc_int_id_contato = " & nCtc_Int_Id_Contato
		Set Sp_Agc =  cConexao.Execute( aSql )
		if Sp_Agc.RecordCount > 0 then
			response.Redirect("../net/msg.asp?Msg=Contato possui agenda !!!")
		end if
		bSql =  "select ctm_int_id_contato_cliente from t_cot_mercadoria where ctm_int_id_contato_cliente = " & nCtc_Int_Id_Contato
		'response.Write(bSql)
		'response.End()
		Set Sp_Cot =  cConexao.Execute( bSql )
		if Sp_Cot.recordcount > 0 then
			response.Redirect("../net/msg.asp?Msg=Contato possui cotação !!!")
		end if
		
		cSql =  "exec sp_exc_Contato '" & nCtc_Int_Id_Contato & "'"
		Set Sp_Contato =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Contato =  Nothing
	    Call Log_Seguranca( "Modulo Cliente", "Excluindo Contato Cliente Nº " & nCtc_Int_Id_Contato )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_contato_cliente.asp?Id_Cliente="&Session("Cad_id_Cliente"))		  
elseif nOpcaoCad_Alt_Exc =  7 then 'Casdatro, Alteração, Exclusão Tipo de Retorno da Agenda
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI10" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Retorno Agenda !!!")
		  end if
 		  cTra_Var_Descricao =  Replace( Request.Form("Tra_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_tipo_retorno_agenda  '0','" & cTra_Var_Descricao & "','" & dData & "','" & dData & "','1'" 
		  Set Sp_tipo_retorno_agenda =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_tipo_retorno_agenda =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Adicionando Tipo de Retorno da Agenda" )
	elseif nOpcao =  2 then 
		  if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI10" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Retorno Agenda !!!")
		  end if
 		  nId_tipo_retorno_agenda =  Request.QueryString("Id_tipo_retorno_agenda") 
  		  cTra_Var_Descricao =  Replace( Request.Form("Tra_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_tipo_retorno_agenda  '" & nId_tipo_retorno_agenda & "','" & cTra_Var_Descricao & "','" & dData & "','" & dData & "','2'" 
		  Set Sp_tipo_retorno_agenda =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_tipo_retorno_agenda =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Alterando Tipo de Retorno da Agenda Nº " & nId_tipo_retorno_agenda )
	 elseif nOpcao =  3 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI10" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Retorno Agenda !!!")
		  end if
		  nId_tipo_retorno_agenda = Request.QueryString("Id_tipo_retorno_agenda")
		  cSql =  "sp_exc_tipo_retorno_agenda '" & nId_tipo_retorno_agenda & "'"
		  Set Sp_tipo_retorno_agenda =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_tipo_retorno_agenda =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Excluindo Tipo de Retorno da Agenda Nº " & nId_tipo_retorno_agenda )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_tipo_retorno_agenda.asp")
elseif nOpcaoCad_Alt_Exc =  8 then 'Casdatro, Alteração, Exclusão Tipo de Atendimento da Agenda
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI09" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Tipo Atendimento Agenda !!!")
		  end if
  		  cTad_Var_Descricao =  Replace( Request.Form("Tad_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_tipo_atendimento '0','" & cTad_Var_Descricao & "','" & dData & "','" & dData & "','1'" 
		  Set Sp_tipo_retorno_agenda =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_tipo_retorno_agenda =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Adicionando Tipo de Atendimento da Agenda" )
	elseif nOpcao =  2 then 
		  if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI09" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Tipo Atendimento Agenda !!!")
		  end if
 		  nId_tipo_atendimento =  Request.QueryString("Id_tipo_atendimento") 
  		  cTad_Var_Descricao =  Replace( Request.Form("Tad_Var_Descricao"), "'", "")
		  dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		  cSql =  "sp_cad_alt_tipo_atendimento  '" & nId_tipo_atendimento & "','" & cTad_Var_Descricao & "','" & dData & "','" & dData & "','2'" 
		  Set Sp_tipo_retorno_agenda =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_tipo_retorno_agenda =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Alterando Tipo de Atendimento Nº " & nId_tipo_atendimento )
	 elseif nOpcao =  3 then
		  if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI09" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Tipo Atendimento Agenda !!!")
		  end if
		  nId_tipo_atendimento = Request.QueryString("Id_tipo_atendimento")
		  cSql =  "sp_exc_tipo_atendimento '" & nId_tipo_atendimento & "'"
		  Set Sp_tipo_retorno_agenda =  cConexao.Execute( cSql )
		  Call TrataErroSistema("")
		  Set Sp_tipo_retorno_agenda =  Nothing
		  Call Log_Seguranca( "Modulo Cliente", "Excluindo Tipo de Atendimento Nº " & nId_tipo_atendimento )
	end if
	Call FechaConexaoSistema()
    response.Redirect("con_tipo_atendimento_agenda.asp")
elseif nOpcaoCad_Alt_Exc =  9 then 'Casdatro, Alteração, Exclusão Agenda
	nOpcao =  Request.QueryString("Opcao2")
	nId_Agenda_Cliente_Atual = Request.QueryString("Id_Agenda_Cliente_Atual")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI08" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo  Agenda !!!")
	    end if
		nCli_Int_Id_Cliente = request.QueryString("Id_Cliente")
		nCtc_Int_Id_Contato =  Replace( Request.Form("Ctc_Int_Id_Contato"), "'", "")
		nTad_Int_Id_Tipo_Atendimento =  Replace( Request.Form("Tad_Int_Id_Tipo_Atendimento"), "'", "")
		nTra_Int_Id_Tipo_Retorno_Agenda =  Replace( Request.Form("Tra_Int_Id_Tipo_Retorno_Agenda"), "'", "")
		nLog_Int_Id_Login =  Session("Id_Login")
		nSts_Int_Id_Status =  Replace( Request.Form("Sts_Int_Id_Status"), "'", "")
		cAgc_Var_Titulo =  Replace( Request.Form("Agc_Var_Titulo"), "'", "")
		cAgc_Var_Texto_Msg = Time() &" - "&Replace( Request.Form("Agc_Var_Texto_Msg"), "'", "")
		  if ( Request.Form("Agc_Char_Aviso") = "T" ) then
			  lAgc_Char_Aviso = "T"
		  else
			  lAgc_Char_Aviso = "F"
		  end if
		  if ( Request.Form("Agc_Char_Grupo") = "T" ) then
			  lAgc_Char_Grupo = "T"
		  else
			  lAgc_Char_Grupo = "F"
		  end if		  
		if request.Form("Agc_Data_Aviso") = "" and request.Form("Sts_Int_Id_Status") = 13 then
		dAgc_Data_Aviso = date()
		else
		dAgc_Data_Aviso =  Request.Form("Agc_Data_Aviso")
		end if
		if dAgc_Data_Aviso <> Empty then
			xData = dAgc_Data_Aviso
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Cadastro Inválida !!!")
			end if
			dAgc_Data_Aviso =  Right("0" & Month( xData ), 2 ) & "/" & Right("0" & Day( xData ), 2 ) & "/" &  Year( xData ) & _
			                   " 9:00"
		end if
		dAgc_Int_Data_Cadastro =  Request.Form("Agc_Int_Data_Cadastro")
		if dAgc_Int_Data_Cadastro <> Empty then
			xData = dAgc_Int_Data_Cadastro
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Cadastro Inválida !!!")
			end if
			dAgc_Int_Data_Cadastro = Year( xData ) & Right("0" & Month( xData ), 2 ) & Right("0" & Day( xData ), 2 )
		end if
		dData =  Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 )
		cSql =  "sp_cad_alt_agenda_cliente '0','" & nCli_Int_Id_Cliente & "','" & nCtc_Int_Id_Contato & "','" & nTad_Int_Id_Tipo_Atendimento
		cSql =  cSql & "','" & nTra_Int_Id_Tipo_Retorno_Agenda & "','" & nLog_Int_Id_Login & "','" & nSts_Int_Id_Status & "','" & cAgc_Var_Titulo
		cSql =  cSql & "','" & cAgc_Var_Texto_Msg & "','" & lAgc_Char_Aviso & "','" & lAgc_Char_Grupo & "','" & dAgc_Data_Aviso & "','" & dAgc_Int_Data_Cadastro		  
		cSql =  cSql & "','" & dData & "','1'" 
		Set Sp_Agenda_Cliente =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Agenda_Cliente =  Nothing

		cSql = "select @@Identity as Numero_atual"
		Set RsSql = cConexao.Execute( cSql )
		if ( RsSql.RecordCount > 0 ) and ( not RsSql.Eof ) then
			Session("Cad_Id_Relacionamento_Agenda") = RsSql("Numero_Atual")
		end if
		RsSql.Close
		Set RsSql = Nothing
		
		if ( Session("Id_CRM_Filtrados") <> "" ) then
			dData =  Year( Date() ) & "/" & Right( "0" & Month( Date() ), 2 ) & "/" & Right( "0" & Day( Date() ), 2 ) & " " & Time()
			cSql = "Update t_crm_filtrados set Crmf_Char_Efetivado = 'T', "
			cSql = cSql & "Crmf_Int_Responsavel = '" & Session("Id_Vendedor") & "', "
			cSql = cSql & "Crmf_Int_Id_Agenda = '" & Session("Cad_Id_Relacionamento_Agenda") & "', "
			cSql = cSql & "Crmf_Date_Conclusao = '" & dData & "' "
			cSql = cSql & "where Crmf_Int_Id_Crm_Filtrados = '" & Session("Id_CRM_Filtrados") & "'"
			Set RsSql = cConexao.Execute( cSql )
			Call TrataErroSistema("")
			RsSql.Close : Set RsSql = Nothing
			
			cSql = "update t_cliente set ven_int_id_vendedor = '" & Session("Id_Vendedor") & "' "
			cSql = cSql & "where cli_int_id_cliente = '" & nCli_Int_Id_Cliente & "'"
			Set RsSql = cConexao.Execute( cSql )
			RsSql.Close : Set RsSql = Nothing
			
			Session("Id_CRM_Filtrados") = ""
		end if
		
			cSql = "select Agc_Int_Id_Agenda_Cliente from t_agenda_cliente where agc_var_titulo ='" & cAgc_Var_Titulo & "' and Cli_Int_Id_Cliente = '" & nCli_Int_Id_Cliente & "'"
			Set RsSql = cConexao.Execute( cSql )
			agenda = RsSql("Agc_Int_Id_Agenda_Cliente")
			RsSql.Close : Set RsSql = Nothing
 		
		Call Log_Seguranca( "Modulo Agenda", "Adicionando Agenda Nº " & agenda)
		nId_lista_grupo = request.Form("lista_grupo")

		if ( ( lAgc_Char_Grupo = "F" ) and ( nId_lista_grupo <> empty ) ) then
			cSql =  "sp_cad_alt_relacionamento_agenda '0','" & nId_lista_grupo & "','" & Session("Cad_Id_Relacionamento_Agenda")
			cSql =  cSql & "','" & dAgc_Int_Data_Cadastro & "','" & dData & "','1'" 
			Set Sp_Relacionamento_Agenda =  cConexao.Execute( cSql )
			Call TrataErroSistema("")
			Set Sp_Relacionamento_Agenda =  Nothing
		end if
		response.Redirect("cad_alt_agenda.asp?Opcao=2&Id_Agenda_Cliente="& agenda & "&Id_Cliente="& request.QueryString("id_cliente"))  
	elseif nOpcao = 2 then
			if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI08" ) ) then
				response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Agenda !!!")
			end if
			nId_Agenda = Request.QueryString("Id_Agenda")
			nId_Status = Request.Form("Sts_Int_Id_Status")
			cSql = "update t_agenda_cliente set sts_int_id_status = '" & nId_Status & "' where agc_int_id_agenda_cliente = '" & nId_Agenda & "'" 
			Set Sp_Agenda_Cliente =  cConexao.Execute( cSql )
			Call TrataErroSistema("")
			Set Sp_Agenda_Cliente =  Nothing
			Call Log_Seguranca( "Modulo Cliente", "Alterando Agenda Nº " & nId_Agenda )
			Call FechaConexaoSistema()
	        response.Redirect("cad_alt_agenda.asp?Opcao=2&Id_Agenda_Cliente="& nId_Agenda & "&Id_Cliente="& request.QueryString("id_cliente"))
	elseif nOpcao =  3 then
		if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI08" ) ) then
		  		response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo  Agenda !!!")
	    end if
		nId_Agenda_Cliente = Request.QueryString("Id_Agenda_Cliente")
		cSql =  "sp_exc_agenda_cliente '" & nId_Agenda_Cliente & "'"
		Set Sp_Agenda_Cliente =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Agenda_Cliente =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Excluindo Tipo de Agenda Nº " & nId_Agenda_Cliente )
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_agenda.asp?Opcao=1&Id_Contato=1&Id_Cliente=" & request.QueryString("Id_Cliente"))
	end if
	Call FechaConexaoSistema()
   if ( (nId_Agenda_Cliente_Atual = nId_Agenda_Cliente) or nId_Agenda_Cliente_Atual = empty) then
     response.Redirect("cad_alt_agenda.asp?Opcao=2&Id_Agenda_Cliente="& Session("Cad_Id_Relacionamento_Agenda"))
   else
	 response.Redirect("cad_alt_agenda.asp?Opcao=2&Id_Agenda_Cliente="& nId_Agenda_Cliente_Atual&"&Id_Contato="&Request.QueryString("Id_Contato"))
   end if
elseif nOpcaoCad_Alt_Exc = 10 then 'Casdatro, Alteração, Exclusão Serviços Mensais
	nOpcao =  Request.QueryString("Opcao2")
	nId_Agenda_Cliente_Atual = Request.QueryString("Id_Agenda_Cliente_Atual")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI13" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Serviços Mensais !!!")
	    end if
		
		nPcs_Int_Id_Produto_Composto_Servico = Replace( Request.Form("Pcs_Int_Id_Produto_Composto_Servico"), "'", "" )
		if ( nPcs_Int_Id_Produto_Composto_Servico <> Empty ) then
			if nPcs_Int_Id_Produto_Composto_Servico = "" then
				Response.Redirect("../net/msg.asp?Msg=Código do Produto Composto Serviço não pode ficar em branco !!!")
			end if
			nPsc_Int_Id_Produto_Servico = 0
		else
			nPsc_Int_Id_Produto_Servico = Replace( Request.Form("Psc_Int_Id_Produto_Servico"), "'", "" )
			if nPsc_Int_Id_Produto_Servico = "" then
				Response.Redirect("../net/msg.asp?Msg=Código do Produto Serviço não pode ficar em branco !!!")
			end if
			nPcs_Int_Id_Produto_Composto_Servico = 0
		end if
		nStatus = 13
        nSmc_Int_Qtd = Replace( Request.Form("Smc_Int_Qtd"), "'", "" )
        nPpg_Int_Id_Planos_Pagamento = Replace( Request.Form("Ppg_Int_Id_Planos_Pagamento"), "'", "" )
        nDtp_Int_Id_Datas_Pagamento = Replace( Request.Form("Dtp_Int_Id_Datas_Pagamento"), "'", "" )

        dSmc_Date_Vencto = Replace( Request.Form("Smc_Date_Vencto"), "'", "" )

		if dSmc_Date_Vencto <> Empty then
			xData = dSmc_Date_Vencto
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Vencimento Invalida !!!")
			end if
			dSmc_Date_Vencto = Year( xData ) & "/" & Right("0" & Month( xData ), 2 ) & "/" & Right("0" & Day( xData ), 2 )
		end if
		
		dSmc_Date_Temp = Replace( Request.Form("Smc_Date_Temp"), "'", "" )

		if dSmc_Date_Temp <> Empty then
			yData = dSmc_Date_Temp
			mData = IsDate( yData )
			if Not mData then
				Response.Redirect("../net/msg.asp?Msg=Data de Vencimento Invalida !!!")
			end if
			dSmc_Date_Temp = Year( yData ) & "/" & Right("0" & Month( yData ), 2 ) & "/" & Right("0" & Day( yData ), 2 )
		end if
		
        nSmc_Dec_Valor = Replace( Replace( Replace( Request.Form("Smc_Dec_Valor2"), "'", "" ), ".", "" ), ",", "." )

		dData =  Year( Date() ) & "/" & Right( "0" & Month( Date() ), 2 ) & "/" & Right( "0" & Day( Date() ), 2 ) & " " & Time()

		cSmc_Char_Valor_Digitado = Request.Form("Smc_Char_Valor_Digitado")
		if ( cSmc_Char_Valor_Digitado = Empty ) then cSmc_Char_Valor_Digitado = "F"
		
		cSql = "exec sp_cad_alt_servicos_mensais '0', '" & Session("Cad_Id_Cliente") & "', '" & nPsc_Int_Id_Produto_Servico & "', '" & nStatus & "', '" 
		cSql = cSql & nSmc_Int_Qtd & "', '" & nPpg_Int_Id_Planos_Pagamento & "', '" & nDtp_Int_Id_Datas_Pagamento & "', '" & dSmc_Date_Vencto & "', '"
		cSql = cSql & nSmc_Dec_Valor & "', '" & nPcs_Int_Id_Produto_Composto_Servico & "', '" & cSmc_Char_Valor_Digitado & "', '" & dData & "', '" & dData & "', '" & dSmc_Date_Temp & "', '1'"
'		response.Write(cSql)
'		response.End()
		
		Set Sp_Servicos_Mensais =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Servicos_Mensais =  Nothing

		Call Log_Seguranca( "Modulo Cliente", "Adicionando Serviços Mensais para o Cliente Nº " & Session("Cad_Id_Cliente") )
		Call FechaConexaoSistema()
		response.Redirect("con_servicos_mensais.asp")
    elseif nOpcao = 2 then
		if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI13" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Serviços Mensais !!!")
	    end if
		nId_Servicos_Mensais = Request.QueryString("Id_Servicos_Mensais")

		nPcs_Int_Id_Produto_Composto_Servico = Replace( Request.Form("Pcs_Int_Id_Produto_Composto_Servico"), "'", "" )
		if ( nPcs_Int_Id_Produto_Composto_Servico <> Empty ) then
			if nPcs_Int_Id_Produto_Composto_Servico = "" then
				Response.Redirect("../net/msg.asp?Msg=Código do Produto Composto Serviço não pode ficar em branco !!!")
			end if
			nPsc_Int_Id_Produto_Servico = 0
			cCaminho = "cad_alt_servicos_mensais_comp"
		else
			nPsc_Int_Id_Produto_Servico = Replace( Request.Form("Psc_Int_Id_Produto_Servico"), "'", "" )
			if nPsc_Int_Id_Produto_Servico = "" then
				Response.Redirect("../net/msg.asp?Msg=Código do Produto Serviço não pode ficar em branco !!!")
			end if
			nPcs_Int_Id_Produto_Composto_Servico = 0
			cCaminho = "cad_alt_servicos_mensais"
		end if
        nSmc_Int_Qtd = Replace( Request.Form("Smc_Int_Qtd"), "'", "" )

		nSts_Int_Id_Status = Replace( Request.Form("Sts_Int_Id_Status"), "'", "" )
        nPpg_Int_Id_Planos_Pagamento = Replace( Request.Form("Ppg_Int_Id_Planos_Pagamento"), "'", "" )
        nDtp_Int_Id_Datas_Pagamento = Replace( Request.Form("Dtp_Int_Id_Datas_Pagamento"), "'", "" )

        dSmc_Date_Vencto = Replace( Request.Form("Smc_Date_Vencto"), "'", "" )
		if dSmc_Date_Vencto <> Empty then
			xData = dSmc_Date_Vencto
			lData = IsDate( xData )
			if Not lData then
				Response.Redirect("../net/msg.asp?Msg=Data de Vencimento Invalida !!!")
			end if
			dSmc_Date_Vencto = Year( xData ) & "/" & Right("0" & Month( xData ), 2 ) & "/" & Right("0" & Day( xData ), 2 )
		end if
		
		dSmc_Date_Temp = Replace( Request.Form("Smc_Date_Temp"), "'", "" )
		if dSmc_Date_Temp <> Empty then
			yData = dSmc_Date_Temp
			mData = IsDate( yData )
			if Not mData then
				Response.Redirect("../net/msg.asp?Msg=Data Temporária Invalida !!!")
			end if
			dSmc_Date_Temp = Year( yData ) & "/" & Right("0" & Month( yData ), 2 ) & "/" & Right("0" & Day( yData ), 2 )
		end if
        nSmc_Dec_Valor = Replace( Replace( Replace( Request.Form("Smc_Dec_Valor2"), "'", "" ), ".", "" ), ",", "." )

		dData =  Year( Date() ) & "/" & Right( "0" & Month( Date() ), 2 ) & "/" & Right( "0" & Day( Date() ), 2 ) & " " & Time()
		
		cSmc_Char_Valor_Digitado = Request.Form("Smc_Char_Valor_Digitado")
		if ( cSmc_Char_Valor_Digitado = Empty ) then cSmc_Char_Valor_Digitado = "F"
		
		cSql = "exec sp_cad_alt_servicos_mensais '" & nId_Servicos_Mensais & "', '" & Session("Cad_Id_Cliente") & "', '" & nPsc_Int_Id_Produto_Servico & "', '" & nSts_Int_Id_Status & "', '" 
		cSql = cSql & nSmc_Int_Qtd & "', '" & nPpg_Int_Id_Planos_Pagamento & "', '" & nDtp_Int_Id_Datas_Pagamento & "', '" & dSmc_Date_Vencto & "', '"
		cSql = cSql & nSmc_Dec_Valor & "', '" & nPcs_Int_Id_Produto_Composto_Servico & "', '" & cSmc_Char_Valor_Digitado & "', '" & dData & "', '" & dData & "', '" & dSmc_Date_Temp & "', '2'"
'		response.Write(cSql)
'		response.End()

		Set Sp_Servicos_Mensais =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Servicos_Mensais =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Alterando Serviços Mensais para o Cliente Nº " & Session("Cad_Id_Cliente") & " Serviço Nº " & nId_Servicos_Mensais )
		Call FechaConexaoSistema()
		response.Redirect(cCaminho&".asp?Opcao=2&Id_Servicos_Mensais="& nId_Servicos_Mensais)
    elseif nOpcao = 3 then
		if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI13" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Serviços Mensais !!!")
	    end if
		nId_Servicos_Mensais = Request.QueryString("Id_Servicos_Mensais")
		
		cSql = "exec sp_exc_servicos_mensais '" & nId_Servicos_Mensais & "'"
		
		Set Sp_Servicos_Mensais =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Servicos_Mensais =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Excluindo Serviços Mensais para o Cliente Nº " & Session("Cad_Id_Cliente") & " Serviço Nº " & nId_Servicos_Mensais )
		Call FechaConexaoSistema()
		response.Redirect("con_servicos_mensais.asp")
	end if
elseif nOpcaoCad_Alt_Exc = 11 then 'Casdatro, Alteração, Exclusão Serviços Mensais ( Composto )
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI13" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Serviços Mensais !!!")
	    end if
		
		nId_Servicos_Mensais = Request.QueryString("Id_Servicos_Mensais")
		nQtd = "1"
		nId_Produto_Servico = Request.QueryString("Id_Produto_Servico")
		nValor = Replace( Request.QueryString("Valor_Produto"), ",", "." )

		dData = Year( Date() ) & "/" & Right("0" & Month( Date() ), 2 ) & "/" & Right("0" & Day( Date() ), 2 ) & " " & Time()
		
		cSql = "Select Ppg_Int_Id_Planos_Pagamento, Dtp_Int_Id_Datas_Pagamento "
		cSql = cSql & "from t_servicos_mensais_cliente "
		cSql = cSql & " where Smc_Int_Servicos_Mensais_Cliente = '" & CInt( nId_Servicos_Mensais ) & "'"
		Set SMC = cConexao.Execute( cSql )
		if ( Not SMC.Eof ) then 
			cPpg_Int_Id_Planos_Pagamento = SMC("Ppg_Int_Id_Planos_Pagamento")
			cDtp_Int_Id_Datas_Pagamento = SMC("Dtp_Int_Id_Datas_Pagamento")
		end if
		SMC.Close : Set SMC = Nothing
			
		cSql = "exec sp_cad_alt_smc_composto '0', '" & nId_Servicos_Mensais & "', '"
		cSql = cSql & nId_Produto_Servico & "', '" & cPpg_Int_Id_Planos_Pagamento & "', '" & cDtp_Int_Id_Datas_Pagamento & "', '" & nQtd & "', '" & dData & "', '" & nValor & "', '"
		cSql = cSql & dData & "', '" & dData & "', '1'"


			
		Set Sp_Smc_Composto =  cConexao.Execute( cSql )
'			Call TrataErroSistema("")
		Set Sp_Smc_Composto =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Adicionando Produto Nº " & nId_Produto_Servico & " no Serviços Mensais para o Cliente Nº " &_
		                                      Session("Cad_Id_Cliente") & " Serviço Nº " & nId_Servicos_Mensais )
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_servicos_mensais_comp.asp?Opcao=2&Id_Servicos_Mensais=" & nId_Servicos_Mensais)
    elseif nOpcao = 2 then
		if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI13" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Serviços Mensais !!!")
	    end if
		
		nId_Servicos_Mensais = Session("Id_Servicos_Mensais_Cliente") 

		dData = Year( Date() ) & "/" & Right("0" & Month( Date() ), 2 ) & "/" & Right("0" & Day( Date() ), 2 ) & " " & Time()
		
		For nPos = 1 To CInt( Request.Form("Pos") )
			
			nId_Smc_Composto = Request.Form("interno_smcc_"&nPos)
			nId_Produto_Servico = Request.Form("codigo_smcc_"&nPos)
			nQtd = TiraSujeira( Request.Form("qtd_smcc_"&nPos) )
			nValor = TiraSujeira ( Request.Form("valor_smcc_"&nPos) )

			nValor = Replace( nValor, ",", "." )
			
			cSql = "exec sp_cad_alt_smc_composto '" & nId_Smc_Composto & "', '" & nId_Servicos_Mensais & "', '"
			cSql = cSql & nId_Produto_Servico & "', '0', '0', '" & nQtd & "', '" & dData & "', '" & nValor & "', '"
			cSql = cSql & dData & "', '" & dData & "', '2'"
			Set Sp_Smc_Composto =  cConexao.Execute( cSql )
'			Call TrataErroSistema("")
			Set Sp_Smc_Composto =  Nothing
			Call Log_Seguranca( "Modulo Cliente", "Alterando Serviços Mensais para o Cliente Nº " & Session("Cad_Id_Cliente") & " Serviço Nº " & nId_Servicos_Mensais )
		Next
'response.End()
		
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_servicos_mensais_comp.asp?Opcao=2&Id_Servicos_Mensais="& nId_Servicos_Mensais)
    elseif nOpcao = 3 then
		if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI13" ) ) then
			response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Serviços Mensais !!!")
	    end if
		nId_Servicos_Mensais = Session("Id_Servicos_Mensais_Cliente") 
		nId_PS_Composto_Produto = Request.QueryString("Id_PS_Composto_Produto")
		
		cSql = "exec sp_exc_smc_composto '" & nId_PS_Composto_Produto & "', '" & nId_Servicos_Mensais & "'"
		
		Set Sp_Servicos_Mensais =  cConexao.Execute( cSql )
		Call TrataErroSistema("")
		Set Sp_Servicos_Mensais =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Excluindo Produto Serviços Mensais para o Cliente Nº " & Session("Cad_Id_Cliente") & " Serviço Nº " & nId_Servicos_Mensais )
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_servicos_mensais_comp.asp?Opcao=2&Id_Servicos_Mensais="& nId_Servicos_Mensais)
	end if
elseif nOpcaoCad_Alt_Exc = 12 then 'Casdatro, Alteração, Exclusão Representante
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		 if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI01" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		  end if
		
		nCli_int_id_cliente = Replace( Request.QueryString("id_cliente"), ",", "." )

		cSql = "exec sp_exc_rep '" & nCli_int_id_cliente & "'"

		Set Sp_Rep =  cConexao.Execute( cSql )
		'Call TrataErroSistema("")
		Set Sp_Rep =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Excluindo representante do cliente Nº " & nCli_int_id_cliente)
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_cliente.asp?Opcao=2&Id_Cliente=" & nCli_int_id_cliente)
	end if
elseif nOpcaoCad_Alt_Exc = 13 then 'Cadastro, Alteração, Exclusão Certificação
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
    if nOpcao =  1 then
		 if ( Not VerificaPermissao( Session("Id_Login"), 1, "CLI14" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		  end if
		nCli_Int_Id_Cliente = request.QueryString("Id_Cliente")
		nCert_Var_Descricao = replace(request.Form("Cert_Var_Descricao"),"'","")
		nFab_Int_Id_Fabricante = request.Form("Fab_Int_Id_Fabricante")
		nCert_Var_Cod_Alternativo = replace(request.Form("Cert_Var_Cod_Alternativo"),"'","")
		nCert_Var_Nivel = replace(request.Form("Cert_Var_Nivel"),"'","")
		if (nFab_Int_Id_Fabricante = 18) then
			if (nCert_Var_Nivel <> "expert" and nCert_Var_Nivel <> "professional" and nCert_Var_Nivel <> "associate") then
				 response.Redirect("../net/msg.asp?Msg=Nivel Watchguard deve ser expert, professional ou associate  !!!")
			end if
		end if
		nCert_Dec_Desconto = replace(request.Form("Cert_Dec_Desconto"),",",".")
		dData = Year( Date() ) & Right("0" & Month( Date() ), 2 ) & Right("0" & Day( Date() ), 2 )
		nCert_Int_Data_Cadastro = dData
		if request.Form("Cert_Int_Data_Certificacao") <> Empty then
			nCert_Int_Data_Certificacao = year(request.Form("Cert_Int_Data_Certificacao")) & right("0" & month(request.Form("Cert_Int_Data_Certificacao")),2) & right("0"&day(request.Form("Cert_Int_Data_Certificacao")),2)
		else
			nCert_Int_Data_Certificacao = ""
		end if
		if request.Form("Cert_Int_Data_Validade") <> empty then
			nCert_Int_Data_Validade = year(request.Form("Cert_Int_Data_Validade")) & right("0" & month(request.Form("Cert_Int_Data_Validade")),2) & right("0"&day(request.Form("Cert_Int_Data_Validade")),2)
		else
			nCert_Int_Data_Validade = ""
		end if
		
		cSql = "exec sp_cad_alt_certificacao '0','" & nCli_Int_Id_Cliente & "','" & nFab_Int_Id_Fabricante & "','" & nCert_Var_Cod_Alternativo & "','" & nCert_Var_Descricao & "','" & nCert_Var_Nivel & "','" & nCert_Int_Data_Certificacao & "','" & nCert_Int_Data_Validade & "','" & nCert_Dec_Desconto & "','" & nCert_Int_Data_Cadastro & "','" & nCert_Int_Data_Cadastro & "','1'"
		'response.Write(csql)
		'response.End()
		Set Sp_Rep =  cConexao.Execute( cSql )
		'Call TrataErroSistema("")
		Set Sp_Rep =  Nothing
		cSql = "select @@Identity as Numero_atual"
		Set RsSql = cConexao.Execute( cSql )
		if ( RsSql.RecordCount > 0 ) and ( not RsSql.Eof ) then
			nCert_Int_Id_Certificacao = RsSql("Numero_Atual")
		end if
		RsSql.Close : Set RsSql = Nothing
		Call Log_Seguranca( "Modulo Cliente", "Adicionando Certificação Nº " & nCert_Int_Id_Certificacao)
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_certificacao.asp?Opcao=2&Id_Certificacao="& nCert_Int_Id_Certificacao &"&Id_Cliente=" & nCli_int_id_cliente)
	elseif nOpcao = 2 then
		if ( Not VerificaPermissao( Session("Id_Login"), 2, "CLI14" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nCert_Int_Id_Certificacao = request.QueryString("Id_Certificacao")
		nCli_Int_Id_Cliente = request.QueryString("Id_Cliente")
		nCert_Var_Descricao = replace(request.Form("Cert_Var_Descricao"),"'","")
		nFab_Int_Id_Fabricante = request.Form("Fab_Int_Id_Fabricante")
		nCert_Var_Cod_Alternativo = replace(request.Form("Cert_Var_Cod_Alternativo"),"'","")
		nCert_Var_Nivel = replace(request.Form("Cert_Var_Nivel"),"'","")
		if (nFab_Int_Id_Fabricante = 18) then
			if (nCert_Var_Nivel <> "expert" and nCert_Var_Nivel <> "professional" and nCert_Var_Nivel <> "associate") then
				 response.Redirect("../net/msg.asp?Msg=Nivel Watchguard deve ser expert, professional ou associate  !!!")
			end if
		end if
		nCert_Dec_Desconto = replace(request.Form("Cert_Dec_Desconto"),",",".")
		if nCert_Dec_Desconto = empty then
			nCert_Dec_Desconto = 0
		end if
		dData = Year( Date() ) & Right("0" & Month( Date() ), 2 ) & Right("0" & Day( Date() ), 2 )
		nCert_Int_Data_Sistema = dData
		if request.Form("Cert_Int_Data_Certificacao") <> Empty then
			nCert_Int_Data_Certificacao = year(request.Form("Cert_Int_Data_Certificacao")) & right("0" & month(request.Form("Cert_Int_Data_Certificacao")),2) & right("0"&day(request.Form("Cert_Int_Data_Certificacao")),2)
		else
			nCert_Int_Data_Certificacao = ""
		end if
		if request.Form("Cert_Int_Data_Validade") <> empty then
			nCert_Int_Data_Validade = year(request.Form("Cert_Int_Data_Validade")) & right("0" & month(request.Form("Cert_Int_Data_Validade")),2) & right("0"&day(request.Form("Cert_Int_Data_Validade")),2)
		else
			nCert_Int_Data_Validade = ""
		end if
		
		cSql = "exec sp_cad_alt_certificacao '" & nCert_Int_Id_Certificacao & "','" & nCli_Int_Id_Cliente & "','" & nFab_Int_Id_Fabricante & "','" & nCert_Var_Cod_Alternativo & "','" & nCert_Var_Descricao & "','" & nCert_Var_Nivel & "','" & nCert_Int_Data_Certificacao & "','" & nCert_Int_Data_Validade & "','" & nCert_Dec_Desconto & "','','" & nCert_Int_Data_Sistema & "','2'"

		Set Sp_Rep =  cConexao.Execute( cSql )
		'Call TrataErroSistema("")
		Set Sp_Rep =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Alterando Certificação Nº " & nCert_Int_Id_Certificacao)
		Call FechaConexaoSistema()
		response.Redirect("cad_alt_certificacao.asp?Opcao=2&Id_Certificacao="& nCert_Int_Id_Certificacao &"&Id_Cliente=" & nCli_int_id_cliente)
	elseif nOpcao = 3 then
		if ( Not VerificaPermissao( Session("Id_Login"), 3, "CLI14" ) ) then
			 response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
		end if
		nCli_int_id_cliente = request.QueryString("Id_Cliente")
		nCert_Int_Id_Certificacao = Request.QueryString("Id_Certificacao")
		cSql = "exec sp_exc_certificacao '" & nCert_Int_Id_Certificacao & "'"
		'response.Write(csql)
		'response.End()
		Set Sp_Rep =  cConexao.Execute( cSql )
		'Call TrataErroSistema("")
		Set Sp_Rep =  Nothing
		Call Log_Seguranca( "Modulo Cliente", "Excluindo Certificação Nº " & nCert_Int_Id_Certificacao)
		Call FechaConexaoSistema()
		response.Redirect("con_certificacao.asp?Reseta_Busca=R&Id_Cliente=" & nCli_int_id_cliente)
	end if
elseif nOpcaoCad_Alt_Exc = 14 then 'Casdatro, Alteração, Exclusão de Permissão Contat Fabricante
	nOpcao =  Request.QueryString("Opcao2")
	Call ConexaoSistema()
    on error resume next 
	if nOpcao = 2 then 
		'if ( Not VerificaPermissao( Session("Id_Login"), 1, "SEG08" ) ) then
			'response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Cadastro de Estados !!!")
	    'end if
		nFab_Int_Id_Fabricante = request.QueryString("Id_Fabricante")
		nCF_Int_Id_Ctc_Fab = request.Querystring("Id_Permissao")
		nCtc_Int_Id_Contato = request.QueryString("Id_Contato")
		nFab_Char_Permisao = request.Form("Fab_Char_Permissao")
		dData = Year( Date() ) & Right( "0" & Month( Date() ), 2 ) & Right( "0" & Day( Date() ), 2 ) 
		
		cSql = "sp_cad_alt_permissao_fabricante_contato  '" & nCtc_Int_Id_Contato & "','" & nCF_Int_Id_Ctc_Fab & "', '" & nFab_Char_Permisao & "','" & dData & "', '2'"
		'response.Write(cSql) 
		'response.End()
		cConexao.Execute( cSql )
		
		Call TrataErroSistema("")
	    Call Log_Seguranca( "Modulo Segurança", "Alterando Permissao Fabricante Contato Nº " & nCF_Int_Id_Ctc_Fab )
		Call FechaConexaoSistema()
    	response.Redirect("cad_alt_contato.asp?Opcao=2&Id_contato=" & nCtc_Int_Id_Contato )
	end if	
end if
%>
