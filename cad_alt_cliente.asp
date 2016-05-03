<!--#include file="../net/conn.net"-->
<!--#include file="../net/utilitarios.net"-->
<!--#include file="../net/util_per.net"-->
<%

Session("QueryString") = ""
if ( Not VerificaPermissaoBD( Session("Id_Login"), 4, "CLI01" ) ) then
	response.Redirect("../net/msg.asp?Msg=Usuário sem permissão no Modulo Cliente !!!")
end if 

Dim aDados( 11,1 )
if ( Request.QueryString("Opcao") = "2" ) then
	Call ConexaoSistema()
if request.QueryString("id_representante") <> empty then
	cConexao.Execute ("delete t_cliente_representante where cli_int_id_cliente = " & request.QueryString("id_cliente"))
	cConexao.Execute ("insert into t_cliente_representante values (" & request.QueryString("id_cliente")& ", " & request.QueryString("id_representante") & ")")
end if
	agSql = " select top 1 ctc_int_id_contato, ctc_int_data_cadastro from t_contato where cli_int_id_cliente = '" & Request.QueryString("Id_Cliente") & "' order by ctc_int_data_cadastro desc "
	
	set rsAg = cConexao.Execute ( agSql )
	dim id_contato
	if ( Not RsAg.Eof ) then 
	id_contato = rsAg("ctc_int_id_contato")
	end if

	cSql = "Select Cli_Int_Id_Cliente, Cli_Var_Razao_Social, Cli_Var_Fantasia, Cli_Char_Tipo_Pessoa, Cli_Var_CNPJ, Sgc_Int_Id_Segmento_Mercado, "
	cSql = cSql & "Cli_Var_IE, Cli_Var_CCM, Cli_Dec_Limite_Credito, Sts_Int_Id_Status, Ven_Int_id_Vendedor, Tpc_Int_Id_Tipo_Cliente, "
	cSql = cSql & "Cli_Char_Ddd, Cli_Var_Telefone, Cli_Var_Ramal, Cli_Var_Fax, Cli_Var_Ramal_Fax, Cli_Var_Home_Page, Cli_Text_Obs, "
	cSql = cSql & "Cli_Text_Obs_Cadastro, Cli_Char_Flag_Cadastro, Cli_Int_Data_Cadastro, Cli_Int_Data_Sistema, Cli_char_flag_sm, Cli_Char_Flag_Representante, Class_Int_Id_Classificacao_Cliente, Pa_int_id_pais "
	cSql = cSql & " from t_cliente "
	cSql = cSql & " where Cli_Int_Id_Cliente = '" & Request.QueryString("Id_Cliente") & "'"
	Set RsCliente = cConexao.Execute( cSql )
	if ( Not RsCliente.Eof ) then 
		Session("Cad_Id_Cliente") = RsCliente("Cli_Int_Id_Cliente")
		nTpc_Int_Id_Tipo_Cliente = RsCliente("Tpc_Int_Id_Tipo_Cliente")
		nSgc_Int_Id_Segmento_Mercado = RsCliente("Sgc_Int_Id_Segmento_Mercado")
		nClass_Int_Id_Classificacao_cliente = RsCliente("Class_Int_Id_Classificacao_Cliente")
		nVen_Int_id_Vendedor = RsCliente("Ven_Int_id_Vendedor") 
		Session("Id_Vendedor_Cliente") = RsCliente("Ven_Int_id_Vendedor") 
		aDados( 11,1 ) = nVen_Int_id_Vendedor
		nSts_Int_Id_Status = RsCliente("Sts_Int_Id_Status")
		Session("Status_Cliente") = nSts_Int_Id_Status
		cCli_Var_Razao_Social = RsCliente("Cli_Var_Razao_Social")
		Session("Cad_Razao_Social") = RsCliente("Cli_Var_Razao_Social")
		cCli_Var_Fantasia = RsCliente("Cli_Var_Fantasia") 
		cCli_Var_CCM = RsCliente("Cli_Var_CCM") 
		cCli_Var_CNPJ = RsCliente("Cli_Var_CNPJ") 
		aDados( 1,1 ) = RsCliente("Cli_Var_CNPJ") 
		cCli_Var_IE = RsCliente("Cli_Var_IE") 
		aDados( 2,1 ) = RsCliente("Cli_Var_IE") 
		cCli_Char_Ddd = RsCliente("Cli_Char_Ddd") 
		aDados( 3,1 ) = RsCliente("Cli_Char_Ddd") 
		cCli_Var_Telefone = RsCliente("Cli_Var_Telefone") 
		aDados( 4,1 ) = RsCliente("Cli_Var_Telefone") 
		cCli_Var_Ramal = RsCliente("Cli_Var_Ramal") 
		aDados( 5,1 ) = RsCliente("Cli_Var_Ramal") 
		cCli_Var_Fax = RsCliente("Cli_Var_Fax") 
		aDados( 6,1 ) = RsCliente("Cli_Var_Fax") 
		cCli_Var_Ramal_Fax = RsCliente("Cli_Var_Ramal_Fax") 
		aDados( 7,1 ) = RsCliente("Cli_Var_Ramal_Fax") 
		nCli_Dec_Limite_Credito = RsCliente("Cli_Dec_Limite_Credito")
		aDados( 8,1 ) = RsCliente("Cli_Dec_Limite_Credito") 
		cTipoPessoa = RsCliente("Cli_Char_Tipo_Pessoa")
		cCli_Char_Tipo_Pessoa = RsCliente("Cli_Char_Tipo_Pessoa") 
		aDados( 9,1 ) = RsCliente("Cli_Char_Tipo_Pessoa") 
		cCli_Var_Home_Page = RsCliente("Cli_Var_Home_Page") 
		aDados( 10,1 ) = RsCliente("Cli_Var_Home_Page") 
		Session("aDados") = aDados
		cCli_Text_Obs = RsCliente("Cli_Text_Obs") 
		cCli_Text_Obs_Cadastro = RsCliente("Cli_Text_Obs_Cadastro") 
		nPa_Int_Id_Pais = RsCliente("Pa_int_id_pais")
		
		dCli_Int_Data_Cadastro = RsCliente("Cli_Int_Data_Cadastro") 
		dCli_Int_Data_Sistema = RsCliente("Cli_Int_Data_Sistema") 
		
		cCli_char_flag_sm = RsCliente("Cli_char_flag_sm") 
		cCli_char_flag_representante = RsCliente("Cli_char_flag_representante")
	end if 
	RsCliente.Close : Set RsCliente = Nothing
	RsAg.Close : Set RsAg = Nothing

	Call FechaConexaoSistema()
else
		nPa_Int_Id_Pais = request.Form("Pa_Int_Id_Pais")
		cTipoPessoa = Request.Form("TipoPessoa")
	Session("Cad_Id_Cliente") = "0"
	nSts_Int_Id_Status = 3
	if ( Request.QueryString("Id_Status") = 8 ) then nSts_Int_Id_Status = CInt( Request.QueryString("Id_Status") )
	if ( Request.QueryString("Id_Status") = 8 ) then cDesativaStatus = "disabled=""disabled"""
	nTpc_Int_Id_Tipo_Cliente = 2
end if
call ConexaoSistema()
jSql = "select log_int_id_login from t_login where ven_int_id_vendedor = (select ven_int_id_vendedor from t_cliente where cli_int_id_cliente ='" & Request.QueryString("Id_Cliente") & "')" 
Set RsValidaVendedor = cConexao.Execute( jSql )
if (not RsValidaVendedor.EOF ) then 
Vvendedor = RsValidaVendedor("log_int_id_login")
else 
Vvendedor = 0
end if
RsValidaVendedor.Close : Set RsValidaVendedor = nothing
call FechaConexaoSistema()

call ConexaoSistema()
bSql = "select ven_int_id_vendedor from t_login where log_int_id_login = " & Session("Id_Login")
Set RsIdVendedor = cConexao.Execute( bSql )
IdVendedor = RsIdVendedor("ven_int_id_vendedor")
RsIdVendedor.Close : Set RsIdVendedor = nothing
call FechaConexaoSistema()
%>
<html>
<head>
<title>ERP</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/estilo.css" rel="stylesheet" type="text/css">
<!--#Include File="../net/javascript.net"-->
<script type="text/javascript" language="JavaScript1.2" src="../net/stm31.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

</head>
<body>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top">
	  <table width="980" border="2" cellpadding="0" cellspacing="0" bordercolor="#006699">
        <tr>
          <td>
		  	<!--#Include File="../net/header.net"-->
		  </td>
        </tr>
        <tr bgcolor="#CCD6E0">
          <td height="23"> 
		  	<!--#Include File="../net/menu.net"-->
		  </td>
        </tr>
        <tr>
          <td height="350" valign="top" bgcolor="#f5f5f5"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="23" class="Texto_Navegacao"><a href="../home.asp">Home</a> &gt;
				  <%if Request.QueryString("tipo") = "forn" then%>
				  <a href="con_fornecedor.asp"> Consulta Fornecedor </a> &gt; Cadastro de Fornecedor
				  <%else%>
				  <a href="busca_cliente.asp?Opcao_Busca=1"> Consulta Cliente </a> &gt; Cadastro de Cliente
				  <%end if%>
				</td>
              </tr>
              <tr> 
                <td height="23" align="center" class="Texto_Titulo">
				<%if Request.QueryString("tipo") = "forn" then%>
				Cadastro de Fornecedor
				<%else%>
				Cadastro de Cliente
				<%end if%>
				</td>
              </tr>
              <tr> 
			  	<%cJavascript = "return valida(this,'')"%>
                <td align="center"><form action="arquivo_utilitarios.asp?Opcao=4&Opcao2=<%=Request.QueryString("Opcao")%>&Id_Cliente=<%=Session("Cad_Id_Cliente")%>&tipo=<%=request.QueryString("Tipo")%><%if ( Request.QueryString("Opcao") = "1" ) then%>&Cli_Char_Tipo_Pessoa=<%=cTipoPessoa%><%end if%>" method="post" name="form1" target="_self" onSubmit="<%=cJavascript%>">
                    <table width="980" border="0" cellpadding="2" cellspacing="2" bordercolor="#FFFFFF" bgcolor="#dcdcdc" class="FundoCadastroAlteracao">
						<%if ( Request.QueryString("Opcao") = "1" ) then
							Call ConexaoSistema()
							dim cep
							cep = request.form("Edc_Var_Cep1") & request.form("Edc_Var_Cep2")
							cSql = "select endereco_logradouro,bairro_descricao, cidade_descricao, uf_sigla from endereco " 
							Csql = Csql & "inner join bairro on endereco.bairro_codigo = bairro.bairro_codigo "
							Csql = Csql & "inner join cidade on cidade.cidade_codigo = bairro.cidade_codigo "
							Csql = Csql & "inner join uf on cidade.uf_codigo = uf.uf_codigo "
							Csql = Csql & "where endereco_cep = '" & cep & "'"
							Set RsEndereco_Cliente = cConexao.Execute( cSql )
						
								if ( Not RsEndereco_Cliente.Eof ) then
								cEdc_Var_Endereco = RsEndereco_Cliente("endereco_logradouro")
								cEdc_Var_Bairro = RsEndereco_Cliente("bairro_descricao")
				 				cEdc_Var_Cep = cep
								cEdc_Var_Cidade = RsEndereco_Cliente("cidade_descricao")
								cEdc_Char_Estado = RsEndereco_Cliente("uf_sigla")
								
								end if
						
							if ( RsEndereco_Cliente.Eof ) then
							jSql="select cidade_descricao, uf_sigla from cidade c inner join uf u on u.uf_codigo = c.uf_codigo "
							jSql = jSql & "where cidade_cep =" & cep
							Set RsCidade_Cliente = cConexao.Execute( jSql )
								if (Not RsCidade_Cliente.Eof) then
									cEdc_Var_Cep = cep
									cEdc_Var_Cidade = RsCidade_Cliente("cidade_descricao")
									cEdc_Char_Estado = RsCidade_Cliente("uf_sigla")
								else 
								cEdc_Var_Endereco = "cep Invalido! verificar cep ou preencher manualmente."
								response.redirect("../cliente/cad_alt_End_Cliente2.asp?Opcao=1&Erro=1&Id_Cliente=" & nId_Cliente)
								end if
							RsCidade_Cliente.Close : Set RsCidade_Cliente = Nothing
							end if
						
						RsEndereco_Cliente.Close : Set RsEndereco_Cliente = Nothing
						
						call fechaconexaosistema()
						end if%>
						<%if ( Request.QueryString("Opcao") = "1" ) then%>
					  <tr> 
                        <td width="133" class="TextoFormulario">
						<%if nPa_Int_Id_Pais = 1 then%>
						Endere&ccedil;o
						<%else%>
						Address
						<%end if%>						</td>
                        <td colspan="2"><input name="Edc_Var_Endereco" type="text" class="TextoCampoFormulario" id="Edc_Var_Endereco" value="<%=cEdc_Var_Endereco%>" size="70" maxlength="100" Opcao="texto" Campo="Endereço"> 
                          <span class="TextoFormulario">
						   &nbsp;<%if nPa_Int_Id_Pais = 1 then%>
								N&uacute;mero 
							<%else%>
								Number
						<%end if%>
						  
						  
                          <input name="Edc_Int_Numero" type="text" class="TextoCampoFormulario" id="Edc_Int_Numero" value="<%=nEdc_Int_Numero%>" size="10" maxlength="10" Opcao="texto" Campo="Numero, caso sem número preencha com zero ou espaço." onKeyPress="return txtBoxFormat(this, '9999999999', event);">
                        </span> </td>
                      </tr>
                      <tr> 
                        <td width="133" class="TextoFormulario">
						  <%if nPa_Int_Id_Pais = 1 then%>
								Complemento
							<%else%>
								Complement
						<%end if%>						</td>
                        <td colspan="2"><input name="Edc_Var_Complemento" type="text" class="TextoCampoFormulario" id="Edc_Var_Complemento" value="<%=cEdc_Var_Complemento%>" size="15" maxlength="20"> 
                          <span class="TextoFormulario">
						   &nbsp;<%if nPa_Int_Id_Pais = 1 then%>
								Bairro
							<%else%>
								District
						<%end if%>
						   
                          <input name="Edc_Var_Bairro" type="text" class="TextoCampoFormulario" id="Edc_Var_Bairro" value="<%=cEdc_Var_Bairro%>" size="27" maxlength="60" Opcao="texto" Campo="Bairro">
                          &nbsp;<%if nPa_Int_Id_Pais = 1 then%>
								Cep
							<%else%>
								Zip Code
						        <%end if%>						 
                          <input name="Edc_Var_Cep" type="text" class="TextoCampoFormulario" id="Edc_Var_Cep" value="<%=cEdc_Var_Cep%>" size="9" maxlength="8" Opcao="numero" Campo="CEP">
                          <a href="http://www.cep.com.br/resultado_endereco.php?txtCEP=<%=cEdc_Var_Cep%>" target="_blank">&nbsp;<%if nPa_Int_Id_Pais = 1 then%>
								Verifica Cep
							<%else%>								 
						<%end if%>
						  </a> </span></td>
                      </tr>
                      <tr> 
                        <td width="133" height="26" class="TextoFormulario">
						<%if nPa_Int_Id_Pais = 1 then%>
								Cidade
							<%else%>
								City
						<%end if%>						</td>
                        <td colspan="2"><input name="Edc_Var_Cidade" type="text" class="TextoCampoFormulario" id="Edc_Var_Cidade" value="<%=cEdc_Var_Cidade%>" size="32" maxlength="60" opcao="texto" campo="Descricao"> 
                          <span class="TextoFormulario">
						  &nbsp;<%if nPa_Int_Id_Pais = 1 then%>
								Estado
							<%else%>
								State
						<%end if%>
						  </span> 
						  <%call conexaosistema()
						  csql = "select * from t_estado where pa_int_id_pais = '"&nPa_Int_Id_Pais&"'"
						  set rs_estado = cconexao.execute(csql)%>
						  <select name="Edc_Char_Estado" class="TextoCampoFormulario" id="Edc_Char_Estado" Opcao="texto" Campo="Estado">
						    <option value="">&nbsp;</option>
							<%do while not rs_estado.eof%>
                            <option value="<%=trim(rs_estado("Est_Char_Sigla"))%>" <%if trim(cEdc_Char_Estado) = trim(rs_estado("Est_Char_Sigla")) then%>selected="selected"<%end if%>><%=rs_estado("Est_Char_Sigla")%> - <%=rs_estado("Est_var_Descricao")%></option>
							<%rs_estado.movenext : loop
							rs_estado.close : set rs_estado = nothing
							call fechaconexaosistema()%>
                          </select>                        </td>
                      </tr>
                      <tr> 
                        <td width="133" class="TextoFormulario">Tipo de Endere&ccedil;o</td>
                        <td class="TextoFormulario" colspan="2">Faturamento</td>
                      </tr>
					  <tr>
						  <td colspan="5" class="TextoFormulario"><hr width="100%" size="1"></td>
					  </tr>
					  <%end if%>
					  <% if Session("Cad_Id_Cliente") <> "0" then%>
                      <tr> 
                        <td height="24" class="TextoFormulario">C&oacute;digo 
                          do Cliente </td>
                        <td class="TextoCampoFormulario"> 
						<%
						if request.QueryString("opcao") = 2 then
						cli_id = request.QueryString("id_cliente")
						%>
						<input type="hidden" value="<%=cli_id%>" name="cli_id" /><%=cli_id%>
						<%end if%></td>
                        <td clospan="2" class="TextoCampoFormulario"></td>
                      </tr>
                      <%end if%>
                      <tr> 
                        <td colspan="4" class="TextoFormulario"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <%if Request.QueryString("Opcao")<>1 then %>
							<tr> 
                              <td width="99" height="18" class="TextoFormulario">Cadastro 
                                em :</td>
                              <td width="122" class="TextoCampoFormulario"><%=MontaData3( dCli_Int_Data_Cadastro )%></td>
                              <td width="103" class="TextoFormulario">Alterado 
                                em :</td>
                              <td class="TextoCampoFormulario"> 
                                <%
							Call ConexaoSistema()
							cSql = "select top 10 Seg_Var_Nome_Login, Seg_Date_horas from t_Seguranca where seg_var_descricao like '% Nº " & Session("Cad_Id_Cliente") & "' "
						    cSql = cSql & "and seg_var_modulo_acessado = 'Modulo Cliente' "
							cSql = cSql & "order by seg_int_id_seguranca desc"
							Set RsSeg = cConexao.Execute( cSql )
							if (RsSeg.Eof ) then
							%>
							<%=("Sem alterações")%>
							<%else%>
                                <select name="select" class="TextoCampoFormulario">
                                  <%Do While ( Not RsSeg.Eof )%>
                                  <option><%=RsSeg("Seg_Var_Nome_Login")%> - <%=RsSeg("Seg_Date_horas")%></option>
                                  <%RsSeg.MoveNext : Loop%>
                                </select> 
                                <%
							end if
							RsSeg.Close : Set RsSeg = Nothing : Call FechaConexaoSistema()
							%>                              </td>
                            </tr>
							<%end if%>
                          </table></td>
                      </tr>
					  <tr> 
                        <td class="TextoFormulario"><input type="hidden" value="<%=request.Form("id_contato")%>" name="id_contato" /></td>
                        <td colspan="2" class="TextoCampoFormulario"></td>
                        <td align="right" class="TextoCampoFormulario"><span class="TextoFormulario">Classificacao</span>
						<select name="Classificacao" class="TextoCampoFormulario" id="select" Opcao="texto" Campo="Classificação">
                            <%
							Call ConexaoSistema()
							cSql = "Select Class_Int_Id_Classificacao_Cliente, Class_Var_Descricao_Cliente from t_classificacao_cliente "
							cSql = cSql & "where class_int_id_status = '1' "
							if request.QueryString("Tipo") = "forn" then
								cSql = cSql & " and Class_Int_Id_Classificacao_Cliente in (4,5,6) "
							end if
							cSql = cSql & "order by Class_Int_Id_Classificacao_Cliente"
							if request.QueryString("Tipo") = "forn" then
								cSql = cSql & " desc "
							end if
							Set RsCliente = cConexao.Execute( cSql )
							if ( Not RsCliente.Eof ) then
							Do While ( Not RsCliente.Eof )
							%>
                            <option value="<%=RsCliente("Class_Int_Id_Classificacao_Cliente")%>" <%if ( nClass_Int_Id_Classificacao_Cliente = RsCliente("Class_Int_Id_Classificacao_Cliente") and request.QueryString("opcao") = 2 ) then%>selected<%end if%>><%=RsCliente("Class_Var_Descricao_Cliente")%></option>
                            <%
							RsCliente.MoveNext
							Loop
							end if
							
							RsCliente.Close : Set RsCliente = nothing : Call FechaConexaoSistema()
							%>
                      </select>                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Tipo Pessoa</td>
                        <td colspan="2" class="TextoCampoFormulario"> 
                          <%if Request.QueryString("Opcao") = "2" then%>
                          <%if cTipoPessoa = "F" then %>
                          Pessoa Física 
                          <%elseif cTipoPessoa = "J" then %>
                          Pessoa Jurídica 
                          <%end if%>
                          <%elseif Request.QueryString("Opcao") = "1" then%>
							<input type="text" maxlength="15" class="TextoCampoFormulario" name="Cli_Char_Tipo_Pessoa" value="<%if cTipoPessoa = "F" then%>Pessoa Física<%else%>Pessoa Jurídica<%end if%>" disabled="disabled">

                          <%end if
						  %>						  </td>
                        <td align="right" class="TextoCampoFormulario"><span class="TextoFormulario">Status</span> 
                          <%if   Request.QueryString("Opcao") = 1 or ( Request.QueryString("Opcao") = 2  and _
						       ( VerificaPermissaoBD( Session("Id_Login"), 1, "CLI06" ) ) or VerificaPermissaoBD( Session("Id_Login"), 2, "CLI06" )  or Vvendedor = 0  or IdVendedor = Session("Id_Vendedor_Cliente")  ) then%>
                          <%if ( cDesativaStatus <> Empty ) then%>
                          <input type="hidden" value="<%=nSts_Int_Id_Status%>" name="Sts_Int_Id_Status"> 
                          <%end if%>
                          <select name="Sts_Int_Id_Status" <%=cDesativaStatus%>  class="TextoCampoFormulario" id="select2" Opcao="select" Campo="Status Cliente">
                          <%
							Call ConexaoSistema()
							cSql = "Select Sts_Int_Id_Status, Sts_Var_Descricao, Sts_Char_Cliente from t_status "
							cSql = cSql & "where Sts_Char_Cliente = " & "'T'" & "order by Sts_Var_Descricao"
							Set RsCliente = cConexao.Execute( cSql )
							if ( Not RsCliente.Eof ) then
							Do While ( Not RsCliente.Eof )
							%>
                            <option value="<%=RsCliente("Sts_Int_Id_Status")%>" <%if ( CInt( nSts_Int_Id_Status ) = RsCliente("Sts_Int_Id_Status") ) then%>selected<%end if%>><%=RsCliente("Sts_Var_Descricao")%></option>
                            <%
							RsCliente.MoveNext
							Loop
							end if
							RsCliente.Close : Set RsCliente = Nothing : Call FechaConexaoSistema()
							%>
                          </select> 
                          <%else%>
                          <%
							Call ConexaoSistema()
							cSql = "Select Sts_Int_Id_Status, Sts_Var_Descricao from t_status where Sts_Int_Id_Status = '" & nSts_Int_Id_Status & "'"
							Set RsStatus = cConexao.Execute( cSql )
							if ( Not RsStatus.Eof ) then
							%>
                          <%if ( VerificaPermissao( Session("Id_Login"), 4, "CLI06" ) ) then%>
                          <%=RsStatus("Sts_Var_Descricao")%> <input type="hidden" name="Sts_Int_Id_Status" value="<%=RsStatus("Sts_Int_Id_Status")%>"> 
                          <%else%>
                          Sem Permissão para visualizar !!! 
                          <%if Request.QueryString("Opcao") = 1 then%>
                          <input type="hidden" name="Sts_Int_Id_Status" value="<%=nSts_Int_Id_Status%>"> 
                          <%end if%>
                          <%end if%>
                          <%end if
							RsStatus.Close : Set RsStatus = nothing : Call FechaConexaoSistema()
							%>
                          <%end if%>
						  </td>
                      </tr>
					  <%if (cTipoPessoa = "J") then
					  		if (request.QueryString("Tipo") <> "forn") then
					  			if ( Request.QueryString("Opcao") = "1" ) then
							%>
						  <tr> 
							<td class="TextoFormulario">
							<%if nPa_Int_Id_Pais = 1 then%>
								CNPJ
							<% else %>
								TAX ID
							<%end if%></td>
							<td width="300" class="TextoFormulario"> <input type="text" maxlength="18" class="TextoCampoFormulario" name="Cli_Var_CNPJ" <%if nPa_Int_Id_Pais = 1 then%>Opcao="cnpj"<%end if%> value="<%=cCli_Var_CNPJ%>"> 
							  <a href="http://www.receita.fazenda.gov.br/PessoaJuridica/CNPJ/cnpjreva/Cnpjreva_Solicitacao.asp" target="_blank">
			<%if nPa_Int_Id_Pais = 1 then %>&nbsp;Verifica CNPJ <% end if %>
						 </a></td>
            <%if nPa_Int_Id_Pais = 1 then %>
						<td width="304" class="TextoFormulario">IE <span class="TextoCampoFormulario">
                          <input name="Cli_Var_Ie2" type="text" class="TextoCampoFormulario" id="Cli_Var_Ie2" value="<%=cCli_Var_IE%>" size="25" maxlength="20">
                        </span></td><%end if%>
                        <td class="TextoCampoFormulario">&nbsp;</td>
                      </tr>

						  <%
						  elseif ( Request.QueryString("Opcao") = "2" ) then%>
						  <tr> 
							<td class="TextoFormulario">
							<%if nPa_Int_Id_Pais = 1 then%>
								CNPJ
							<% else %>
								TAX ID
							<%end if%></td>
							<td width="300" class="TextoCampoFormulario"><%=cCli_Var_CNPJ%>
							<input type="hidden" name="cCli_Var_CNPJ" value="<%=cCli_Var_CNPJ%>"></td>
							<%if nPa_Int_Id_Pais = 1 then%>
							<td class="TextoFormulario">IE <span class="TextoCampoFormulario">
							  <input name="Cli_Var_Ie" type="text" class="TextoCampoFormulario" id="Cli_Var_Ie" value="<%=cCli_Var_IE%>" size="25" maxlength="20">  
		
		  
							</span></td><%end if%>
							<td class="TextoCampoFormulario">&nbsp;</td>
						  </tr>
						
						
                      	<%end if%>
					<%end if%>
                      <%elseif (cTipoPessoa = "F") then 
		    		 	 if ( Request.QueryString("Opcao") = "1" ) then
					  		if nPa_Int_Id_Pais = 1 then%>
  
	                  <tr> 
	  
                        <td class="TextoFormulario">CPF </td>
                        <td width="300" class="TextoCampoFormulario"> <input type="text" maxlength="18" class="TextoCampoFormulario" name="Cli_Var_CNPJ" Campo="CPF" Opcao="cpf" value="<%=cCli_Var_CNPJ%>"></td>
                        <td class="TextoFormulario">RG <span class="TextoCampoFormulario">
                          <input name="Cli_Var_IE" type="text" class="TextoCampoFormulario" id="Cli_Var_IE" value="<%=cCli_Var_IE%>" size="25" maxlength="20">
                        </span></td>
                        <td class="TextoCampoFormulario">&nbsp;</td>
                      </tr> 
           
					  	  <%end if
					  	elseif ( Request.QueryString("Opcao") = "2" ) then%>
						  <tr> 
							<td class="TextoFormulario">CPF</td>
							<td width="300" class="TextoCampoFormulario"><%=cCli_Var_CNPJ%>
							<input type="hidden" name="cCli_Var_CNPJ" value="<%=cCli_Var_CNPJ%>"></td>
							<td class="TextoFormulario">RG <span class="TextoCampoFormulario"><%=cCli_Var_IE%></span></td>
							<td class="TextoCampoFormulario">&nbsp;</td>
						  </tr> 
					  	<%end if%>
                      <%end if%>
                      <tr> 
                        <td width="133" class="TextoFormulario">Raz&atilde;o Social</td>
                        <td colspan="3" class="TextoFormulario"><input name="Cli_Var_Razao_Social" type="text" c class="TextoCampoFormulario" id="Cli_Var_Razao_Social" value="<%=cCli_Var_Razao_Social%>" size="50" maxlength="100" Opcao="texto" Campo="Razão Social"><a href="busca_cliente.asp?Opcao_Busca=1" target="_self"><img src="../img_sistema/botao/botao_lupa.gif" width="16" height="16" border="0"></a>
                          <a href="busca_cliente.asp?Opcao_Busca=1" target="_self">Procura Cliente</a></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Nome Fantasia</td>
                        <td class="TextoCampoFormulario"> <input name="Cli_Var_Fantasia" type="text" c class="TextoCampoFormulario" id="Cli_Var_Fantasia" value="<%=cCli_Var_Fantasia%>" size="50" maxlength="100" Opcao="texto" Campo="Nome Fantasia">						</td>
						<td align="right" class="TextoFormulario">Representante</td>
                        <td class="TextoFormulario">
							<input type="checkbox" name="cli_char_flag_representante" value="T" <%if ( cCli_char_flag_representante = "T" ) then%>checked<%end if%>>						</td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Tipo Cliente</td>
                        <td colspan="2" class="TextoFormulario"> 
                          <%if ( VerificaPermissaoBD( Session("Id_Login"), 1, "CLI12" ) or VerificaPermissaoBD( Session("Id_Login"), 2, "CLI12" ) or ( Vvendedor = 0 ) ) then%>
                          <select name="Tpc_Int_id_Tipo_Cliente"  class="TextoCampoFormulario" id="select4" Opcao="select" Campo="Tipo Cliente">
                            <option value="" <%if ( nTpc_Int_id_Tipo_Cliente = Empty ) then%>selected<%end if%>>&nbsp;</option>
                            <%
							Call ConexaoSistema()
							cSql = "Select Tpc_Int_id_Tipo_Cliente, Tpc_Var_Descricao from t_tipo_Cliente order by Tpc_Var_Descricao"
							Set RsTipo_Cliente = cConexao.Execute( cSql )
							if ( Not RsTipo_Cliente.Eof ) then
							Do While ( Not RsTipo_Cliente.Eof )
							%>
                            <option value="<%=RsTipo_Cliente("Tpc_Int_id_Tipo_Cliente")%>" <%if ( nTpc_Int_id_Tipo_Cliente = RsTipo_Cliente("Tpc_Int_id_Tipo_Cliente") ) then%>selected<%end if%>><%=RsTipo_Cliente("Tpc_Var_Descricao")%></option>
                            <%
							RsTipo_Cliente.MoveNext
							Loop
							end if
							RsTipo_Cliente.Close : Set RsTipo_Cliente = nothing : Call FechaConexaoSistema()
							%>
                          </select> 
                          <%elseif ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI12" ) ) then%>
                          <%
						  Call ConexaoSistema()
						  cSql = "Select Tpc_Int_id_Tipo_Cliente, Tpc_Var_Descricao from t_tipo_Cliente where Tpc_Int_id_Tipo_Cliente = '" & nTpc_Int_id_Tipo_Cliente & "'"
						  Set RsTipo_Cliente = cConexao.Execute( cSql )
						  if ( Not RsTipo_Cliente.Eof ) then
						  %>
                          <%=RsTipo_Cliente("Tpc_Var_Descricao")%> 
                          <%
						  end if
						  RsTipo_Cliente.Close : Set RsTipo_Cliente = nothing : Call FechaConexaoSistema()
						  %>
                          <input type="hidden" name="Tpc_Int_id_Tipo_Cliente" value="<%=nTpc_Int_id_Tipo_Cliente%>"> 
                          <%else%>
                          Sem Permissão para visualizar !!! 
                          <input type="hidden" name="Tpc_Int_id_Tipo_Cliente" value="<%=nTpc_Int_id_Tipo_Cliente%>"> 
                          <%end if%></td>
                        <td width="217" class="TextoFormulario">&nbsp;</td>
                      </tr>
					  <%if  Request.QueryString("Opcao") <> 1 then%>
					  <tr> 
					  <% if ( Request.QueryString("Opcao") = "2" ) then 
						  Call ConexaoSistema()  
						  set rs = cConexao.execute ("select cli_var_razao_social from t_cliente_representante as tcr inner join t_cliente as tc on (tcr.cli_int_id_representante = tc.cli_int_id_cliente) where tcr.cli_int_id_cliente = " & request.QueryString("id_cliente"))
						  if not rs.eof then
							fantasia = rs("cli_var_razao_social")
						  end if
						  end if%>  
                        <td class="TextoFormulario">Representante<br> </td>
                        <td colspan="2" class="TextoFormulario"><input name="Representante" type="text" class="TextoCampoFormulario" id="Representante" value="<%=fantasia%>" size="50" maxlength="100" disabled="disabled">
						<%if ((nSts_Int_Id_Status = 6 and Session("Id_Login_Master")) or (nSts_Int_Id_Status <> 6)) and (( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) ) then%>
						<%if rs.recordcount <> 0 then%>
						<a href="../cliente/arquivo_utilitarios.asp?Opcao=12&Opcao2=1&Id_Cliente=<%=request.QueryString("id_cliente")%>"><img src="../img_sistema/botao/botao_exclui.gif" width="17" height="17" border="0"></a>
						<%end if%>
                        <a href="javascript:MM_openBrWindow('busca_representante.asp?Opcao_Busca=14&Id_Cliente=<%=request.QueryString("id_cliente")%>','AlteraSenhaUsuarioHome','resizable=yes,width=800,height=400')"><img src="../img_sistema/botao/botao_lupa.gif" width="16" height="16" border="0"> Procura Representante </a></td><%end if%>
						<td class="TextoFormulario">						</td>
					  </tr>
					  <%end if%>
                      <tr> 
					  <%
					  call conexaosistema()
					  set rs2 = cConexao.execute ("Select ven_var_nome from t_cliente_vendedor as tcv inner join t_vendedor as vendedor on (tcv.ven_int_id_vendedor = vendedor.ven_int_id_vendedor) where tcv.cli_int_id_cliente = '" & request.QueryString("Id_Cliente") & "'")%>
                        <td class="TextoFormulario">Vendedor <%if rs2.recordCount > 0 then%><a href="cad_alt_cliente_vendedor.asp?Opcao=2&Id_Cliente=<%=request.QueryString("Id_Cliente")%>">(<%=rs2.recordCount%>)</a><%end if : call fechaconexaosistema()%></td>
                        <td colspan="3" class="TextoCampoFormulario"> 
                          <% if  Request.QueryString("Opcao") = 1  or (  Request.QueryString("Opcao") = 2  and  VerificaPermissaoBD( Session("Id_Login"), 2, "CLI04" ) ) or Vvendedor = 0  or  IdVendedor = Session("Id_Vendedor_Cliente")  then					  
						  %>
                          <select name="Ven_Int_Id_Vendedor" class="TextoCampoFormulario" id="select" Opcao="texto" Campo="Vendedor">
                            <option value="" <%if ( nVen_Int_id_Vendedor = Empty ) then%>selected<%end if%>>&nbsp;</option>
                            <%
							Call ConexaoSistema()
							cSql = "Select Ven_Int_Id_Vendedor, Ven_Var_Nome from t_vendedor "
'							if ( Request.QueryString("Opcao") = 1 ) then
								cSql = cSql & "where sts_int_id_status = '1' "
'							end if
							cSql = cSql & "order by Ven_Var_Nome"
							Set RsCliente = cConexao.Execute( cSql )
							if ( Not RsCliente.Eof ) then
							Do While ( Not RsCliente.Eof )
							%>
                            <option value="<%=RsCliente("Ven_Int_Id_Vendedor")%>" <%if ( nVen_Int_id_Vendedor = RsCliente("Ven_Int_Id_Vendedor") and request.QueryString("opcao") = 2 ) or (request.QueryString("opcao") = 1 and Session("Id_Vendedor") = RsCliente("Ven_Int_Id_Vendedor")) then%>selected<%end if%>><%=RsCliente("Ven_Var_Nome")%></option>
                            <%
							RsCliente.MoveNext
							Loop
							end if
							
							RsCliente.Close : Set RsCliente = nothing : Call FechaConexaoSistema()
							%>
                          </select> 
                          <%else%>
                          <%
							Call ConexaoSistema()
							cSql = "Select Ven_Int_Id_Vendedor, Ven_Var_Nome from t_vendedor where Ven_Int_Id_Vendedor = '" & nVen_Int_id_Vendedor & "'"
							Set RsVendedor = cConexao.Execute( cSql )
							if ( Not RsVendedor.Eof ) then
							%>
                          <%if ( VerificaPermissao( Session("Id_Login"), 4, "CLI04" ) ) then%>
                          <%=RsVendedor("Ven_Var_Nome")%> 
                          <%else%>
                          Sem Permissão para visualizar !!! 
                          <%end if%>
                          <%
							RsVendedor.MoveNext
							end if
							RsVendedor.Close : Set RsVendedor = nothing : Call FechaConexaoSistema()
							%>
                          <%end if%>
						  
						  <% if Request.QueryString("Opcao") = 2 then 
						  if ((nSts_Int_Id_Status = 6 and Session("Id_Login_Master")) or (nSts_Int_Id_Status <> 6)) and (session("id_vendedor") = nven_int_id_vendedor or session("Id_Login_Master") ) then %>
						  <span class="TextoFormulario"><a href="cad_alt_cliente_vendedor.asp?Id_cliente=<%=request.QueryString("Id_cliente")%>" target="_self"><img src="../img_sistema/botao/botao_lupa.gif" width="16" height="16" border="0"> Adicionar Novo Vendedor</a></span>		<%end if%>
						  <%end if%></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Segmento Mercado</td>
                        <td colspan="3" class="TextoCampoFormulario"><select name="Sgc_Int_Id_Segmento_Mercado"  class="TextoCampoFormulario" id="select" Opcao="select" Campo="Segmento Mercado">
                            <option value="" <%if ( nSgc_Int_Id_Segmento_Mercado = Empty ) then%>selected<%end if%>>&nbsp;</option>
                            <%
							Call ConexaoSistema()
							cSql = "Select Sgc_Int_Id_Segmento_Mercado, Sgc_Var_Descricao from t_segmento_mercado order by Sgc_Var_Descricao"
							Set RsSegmento = cConexao.Execute( cSql )
							if ( Not RsSegmento.Eof ) then
							Do While ( Not RsSegmento.Eof )
							%>
                            <option value="<%=RsSegmento("Sgc_Int_Id_Segmento_Mercado")%>" <%if ( nSgc_Int_Id_Segmento_Mercado = RsSegmento("Sgc_Int_Id_Segmento_Mercado") ) then%>selected<%end if%>><%=RsSegmento("Sgc_Var_Descricao")%></option>
                            <%
							RsSegmento.MoveNext
							Loop
							end if
							RsSegmento.Close : Set RsSegmento = nothing : Call FechaConexaoSistema()
							%>
                          </select> </td>
                      </tr>
                      
					  <%if nPa_Int_Id_Pais = 1 then%>
					<tr> 
                       <td class="TextoFormulario">CCM<br> </td>
                        <td colspan="3" class="TextoFormulario"><input name="Cli_Var_CCM" type="text" class="TextoCampoFormulario" id="Cli_Var_CCM" value="<%=cCli_Var_CCM%>" size="25" maxlength="20"></td>
                      </tr>
                    
							<%else%>
					
						<%end if%>		
					  
       <%if nPa_Int_Id_Pais = 1 then%>
					    <tr> 
					    <td class="TextoFormulario">Limite de Cr&eacute;dito<br>                        </td>
                        <td colspan="3" class="TextoFormulario"> 
                          <%if ( Request.QueryString("Opcao") = 1 ) or ( ( Request.QueryString("Opcao") = 2 ) and ( VerificaPermissaoBD( Session("Id_Login"), 2, "CLI05" ) ) or ( Vvendedor = 0 )  ) then%>
                          <input name="Cli_Dec_Limite_Credito" type="text" class="TextoCampoFormulario" id="Cli_Dec_Limite_Credito" value="<%=nCli_Dec_Limite_Credito%>" size="25" maxlength="12"></td>
						  
                        <%else%>
                        <%if ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI05" ) ) then%>
                        <%=nCli_Dec_Limite_Credito%> 
                        <input name="Cli_Dec_Limite_Credito" type="hidden" id="Cli_Dec_Limite_Credito" value="<%=nCli_Dec_Limite_Credito%>">
                        <%else%>
                        <input name="Cli_Dec_Limite_Credito" type="hidden" id="Cli_Dec_Limite_Credito" value="<%=nCli_Dec_Limite_Credito%>">
                        Sem Permissão para visualizar !!! 
                        <%end if%>
                        <%end if%>
                      </tr>
	<%else%><%end if%> 				  
					  
					  <tr><td class="TextoFormulario">País</td><td class="TextoCampoFormulario">
					  <%call conexaosistema()
						set rs = cconexao.execute("Select Pa_Var_Descricao, Pa_Var_DDI, Pa_Int_Id_Pais from t_pais where Pa_Int_Id_Pais = '"&nPa_Int_id_pais&"'")
						if not rs.eof then
							DDI = rs("Pa_Var_DDI")
							nPa_Int_Id_Pais = rs("Pa_Int_Id_Pais")
							nPa_Var_Descricao = rs("Pa_Var_Descricao")
						end if
						%>
					  
					  <select class="TextoCampoFormulario" name="Pa_Int_Id_Pais" id="select" Opcao="select" Campo="País" value="<%=nPa_Int_Id_Pais%>">
						<%
						do while not rs.eof%>
						<option value="<%=nPa_Int_Id_Pais%>" <%if rs("Pa_Int_Id_Pais") = nPa_Int_Id_Pais then%>selected="selected"<%end if%>><%=rs("Pa_Var_Descricao")%></option>
						<%rs.movenext
						loop
						rs.close : set rs = nothing
						call fechaconexaosistema()%>
					  </select></td></tr>
                      <%if ( Request.QueryString("Opcao") = 1 ) or ( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) then%>
                      <tr> 
                        <td colspan="15" class="TextoFormulario"><table width="700" border="0" cellspacing="0" cellpadding="0">
                            <tr>
								<td class="TextoFormulario">DDI</td>
								<td class="TextoCampoFormulario"><%=DDI%></td>
							  <td class="TextoFormulario">DDD</td>
							  <td class="TextoCampoFormulario"><input name="Cli_Char_DDD" type="text" class="TextoCampoFormulario" id="Cli_Char_DDD" value="<%=cCli_Char_DDD%>" size="4" maxlength="3" Opcao="texto" Campo="DDD"></td>
                              <td class="TextoFormulario">Telefone</td>
							  <td class="TextoCampoFormulario"><input name="Cli_Var_Telefone" type="text" class="TextoCampoFormulario" id="Cli_Var_Telefone" value="<%=cCli_Var_Telefone%>" size="12" maxlength="9" Opcao="texto" Campo="Telefone" onKeyPress="return txtBoxFormat(this, '9999-9999', event);"></td>
                              <td class="TextoFormulario">Ramal</td> 
                              <td class="TextoCampoFormulario"><input name="Cli_Var_Ramal" type="text" class="TextoCampoFormulario" id="Cli_Var_Ramal" value="<%=cCli_Var_Ramal%>" size="4" maxlength="5"></td>
                              <td class="TextoFormulario">Fax</td> 
                              <td class="TextoCampoFormulario"><input name="Cli_Var_Fax" type="text" class="TextoCampoFormulario" id="Cli_Var_Fax" value="<%=cCli_Var_Fax%>" size="12" maxlength="9" onKeyPress="return txtBoxFormat(this, '9999-9999', event);">                              </td>
								<td class="TextoFormulario">Ramal</td> 
                                <td class="TextoCampoFormulario"><input name="Cli_Var_Ramal_Fax" type="text" class="TextoCampoFormulario" id="Cli_Var_Ramal_Fax" value="<%=cCli_Var_Ramal_Fax%>" size="4" maxlength="5">                              </td>
                            </tr>
                          </table></td>
                      </tr>
                      <%elseif ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) or ( Vvendedor = 0 ) ) then%>
                      <%if ( VerificaPermissaoBD( Session("Id_Login"), 2, "CLI07" ) or ( Vvendedor = 0 ) ) then%>
                      <tr> 
                        <td colspan="15" class="TextoFormulario"><table width="700" border="0" cellspacing="0" cellpadding="0">
                            <tr>
								<td class="TextoFormulario">DDI</td>
								<td class="TextoCampoFormulario"><%=DDI%></td>
							  <td class="TextoFormulario">DDD</td>
							  <td><input name="Cli_Char_DDD" type="text" class="TextoCampoFormulario" id="Cli_Char_DDD" value="<%=cCli_Char_DDD%>" size="4" maxlength="3" Opcao="texto" Campo="DDD"></td>
                              <td class="TextoFormulario">Telefone</td>
							  <td class="TextoCampoFormulario"><input name="Cli_Var_Telefone" type="text" class="TextoCampoFormulario" id="Cli_Var_Telefone" value="<%=cCli_Var_Telefone%>" size="12" maxlength="9" Opcao="texto" Campo="Telefone" onKeyPress="return txtBoxFormat(this, '9999-9999', event);"></td>
                              <td class="TextoFormulario">Ramal</td>
                              <td class="TextoCampoFormulario"><input name="Cli_Var_Ramal" type="text" class="TextoCampoFormulario" id="Cli_Var_Ramal" value="<%=cCli_Var_Ramal%>" size="4" maxlength="5"></td>
                              <td class="TextoFormulario">Fax</td> 
                              <td class="TextoCampoFormulario"><input name="Cli_Var_Fax" type="text" class="TextoCampoFormulario" id="Cli_Var_Fax" value="<%=cCli_Var_Fax%>" size="12" maxlength="9" onKeyPress="return txtBoxFormat(this, '9999-9999', event);">                              </td>
                              <td class="TextoFormulario">Ramal</td>
                              <td class="TextoCampoFormulario"><input name="Cli_Var_Ramal_Fax" type="text" class="TextoCampoFormulario" id="Cli_Var_Ramal_Fax" value="<%=cCli_Var_Ramal_Fax%>" size="4" maxlength="5">                              </td>
                            </tr>
                          </table></td>
                      </tr>
                      <%else%>
                      <tr> 
                        <td colspan="15" class="TextoFormulario"><table width="700" border="0" cellspacing="0" cellpadding="0">
                            <tr>
								<td class="TextoFormulario">DDI</td>
								<td class="TextoCampoFormulario"><%=DDI%></td>
							  <td class="TextoFormulario">DDD</td>
                              <td class="TextoFormulario"><%=cCli_Char_DDD%></td>
                              <td class="TextoFormulario">Telefone</td>
							  <td class="TextoCampoFormulario"><%=cCli_Var_Telefone%>                              </td>
                              <td class="TextoFormulario">Ramal</td>
							  <td class="TextoCampoFormulario"><%=cCli_Var_Ramal%>                              </td>
                              <td class="TextoFormulario">Fax</td>
							  <td class="TextoCampoFormulario"><%=cCli_Var_Fax%>                              </td>
                              <td class="TextoFormulario">Ramal</td>
							  <td class="TextoCampoFormulario"><%=cCli_Var_Ramal_Fax%>                              </td>
                            </tr>
                          </table></td>
                      </tr>
                      <%end if%>
                      <%end if%>
                      <tr> 
                        <td class="TextoFormulario"> Home Page</td>
                        <td colspan="3"><input name="Cli_Var_Home_Page" type="text" class="TextoCampoFormulario" id="Cli_Var_Home_Page" value="<%=cCli_Var_Home_Page%>" size="50" maxlength="256"><a href="http://<%=cCli_Var_Home_Page%>" target="_blank"><img src="../img_sistema/botao/botao_site.gif" width="20" height="21" border="0"></a></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Servi&ccedil;o Mensal</td>
                        <td class="TextoFormulario"><input type="checkbox" name="cli_char_flag_sm" value="T" <%if ( cCli_char_flag_sm = "T" ) then%>checked<%end if%>></td>
						 <td align="left" class="TextoFormulario">
						 <%if ( request.QueryString("Opcao") = 2 ) and ( ( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or (Session("Id_Login") = 16) or (Session("Id_Login") = 50) or (Session("Id_Login") = 45) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or ( VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor"))) ) )then%>
							 <a href="../vendas_servicos/con_ticket.asp?Id_Cliente=<%=cli_id%>"><img border="0" src="../img_sistema/botao/botao_altera3.gif"> Ticket</a>
						     <%end if%>						 </td>
                        <td class="TextoFormulario">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Obs</td>
                        <td colspan="3"><textarea name="Cli_Text_Obs" cols="70" rows="4" wrap="VIRTUAL" class="TextoCampoFormulario" id="Cli_Text_Obs" c="c" opcao="texto" campo="Comiss&atilde;o"><%=cCli_Text_Obs%></textarea></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Obs Cadastro</td>
                        <td colspan="3"><textarea name="Cli_Text_Obs_Cadastro" cols="70" rows="4" wrap="VIRTUAL" class="TextoCampoFormulario" id="Cli_Text_Obs_Cadastro"><%=cCli_Text_Obs_Cadastro%></textarea>                </td>
                      </tr>
					     <%if ( Request.QueryString("Opcao") = "2" ) then%>
					  <tr> 
                        <td colspan="4" align="center" class="TextoFormulario">
						<%if ((nSts_Int_Id_Status = 6 and Session("Id_Login_Master")) or (nSts_Int_Id_Status <> 6)) and (( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) ) then%>
							<table width="800" border="0" cellspacing="0" cellpadding="0">
                            	<tr> 
                              		<%Call ConexaoSistema()
									set rs_cert = cconexao.execute("select cert_int_id_certificacao from t_certificacao where cli_int_id_cliente = '"& Session("Cad_Id_Cliente") &"'")
									%>
									<td width="198" class="TextoFormulario"><a href="con_certificacao.asp?Reseta_Busca=R&Id_Cliente=<%=Session("Cad_Id_Cliente")%>"><img src="../img_sistema/botao/botao_clausulas.gif" border="0">Certifica&ccedil;&otilde;es (<%response.Write(rs_cert.recordcount)%>)</a></td><%Call Fechaconexaosistema()%>
                              		<td width="198" class="TextoFormulario">&nbsp; 
                               			<a href="con_End_Cliente.asp?Id_Cliente=<%=Session("Cad_Id_Cliente")%>"><img src="../img_sistema/botao/botao_cliente_endereco.gif" width="24" height="28" border="0">Endere&ccedil;o
						<%Call ConexaoSistema()
						Set RsEn = cConexao.Execute("Select edc_int_id_end_cliente from t_end_cliente where Cli_Int_Id_Cliente = '"&Request.QueryString("Id_Cliente")&"'")%>
					  	(<%response.Write(RsEn.recordcount)%>)						</a>									</td>
                              	<td width="232" valign="middle" class="TextoFormulario"> 
                                	<a href="con_contato_cliente.asp?Id_Cliente=<%=Session("Cad_Id_Cliente")%>"><img src="../img_sistema/botao/botao_cliente_contato.gif" width="28" height="28" border="0">Contato
						<%Call ConexaoSistema()
						Set RsCt = cConexao.Execute("select ctc.Ctc_Int_Id_Contato from t_Contato as ctc inner join t_cliente_contato as tcc on (ctc.ctc_int_id_contato = tcc.ctc_int_id_contato) where tcc.Cli_Int_Id_Cliente = '"&Request.QueryString("Id_Cliente")&"'")%>
					  	(<%response.Write(RsCt.recordcount)%>)						</a>						</td>
                              	<td width="270" valign="middle" class="TextoFormulario"> 
                                	<a href="../cliente/cad_Alt_agenda.asp?Opcao=1&Id_Contato=<%=id_contato%>&Id_Cliente=<%=Session("Cad_Id_Cliente")%>"><img src="../img_sistema/botao/botao_agenda.gif" width="28" height="28" border="0">Agenda
						<%Call ConexaoSistema()
						Set RsAg = cConexao.Execute("Select agc_int_id_agenda_cliente from t_agenda_cliente where Cli_Int_Id_Cliente = '"&Request.QueryString("Id_Cliente")&"'")%>
					  	(<%response.Write(RsAg.recordcount)%>)						</a>							   </td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td colspan="4" align="center" class="TextoFormulario">
							<table width="800" border="0" cellspacing="0" cellpadding="0">
                            	<tr> 
                              		<td width="198" class="TextoFormulario" align="left">&nbsp; 
                               			 <a href="../pesquisas/con_pesquisa.asp?Id_Cliente=<%=request.QueryString("id_cliente")%>"><img src="../img_sistema/botao/botao_pesquisa3.png" border="0" width="28" height="28">Pesquisas
								<%Call ConexaoSistema()
								Set RsPq = cConexao.Execute("Select pesq01_int_id_pesq01 from t_pesquisa_01 where cli_int_id_cliente = '" & request.QueryString("id_cliente") & "'")
								Set RsPq2 = cConexao.Execute("Select pesq02_int_id_pesq02 from t_pesquisa_02 where cli_int_id_cliente = '" & request.QueryString("id_cliente") & "'")
								Set RsPq3 = cConexao.Execute("Select pesq03_int_id_pesq03 from t_pesquisa_03 where cli_int_id_cliente = '" & request.QueryString("id_cliente") & "'")
								Set RsPq4 = cConexao.Execute("Select pesq04_int_id_pesq04 from t_pesquisa_04 where cli_int_id_cliente = '" & request.QueryString("id_cliente") & "'")%>
					  			(<%response.Write(RsPq.recordcount + RsPq2.recordcount + RsPq3.recordcount + RsPq4.recordcount)%>)							  </a>									</td>
                              	<td width="232" valign="middle" class="TextoFormulario"> 
                                	<%if ((nSts_Int_Id_Status = 6 and Session("Id_Login_Master")) or (nSts_Int_Id_Status <> 6)) and (( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) ) then%>
                                	<a href="con_posicao_cr_cliente.asp?Opcao_Visualizacao=T"> <img src="../img_sistema/botao/botao_posicao_financeira.gif" width="21" height="21" border="0"> 
                                	Posi&ccedil;&atilde;o Financeira
									<%Call ConexaoSistema()
									Set RsPf = cConexao.Execute("Select Crb_Int_Id_Contas_Receber from t_contas_receber as CR inner join t_nota_fiscal as NF on ( CR.Nfl_Int_Id_Nota_Fiscal = NF.Nfl_Int_Id_Nota_Fiscal ) where CR.Cli_Int_Id_Cliente = '"&Request.QueryString("Id_Cliente")&"' union Select Crb_Int_Id_Contas_Receber from t_contas_receber as CR inner join t_nota_fiscal as NF on ( CR.Nfl_Int_Id_Nota_Fiscal = NF.Nfl_Int_Id_Nota_Fiscal ) where NF.Nfl_Int_Revendedor = '"&Session("Cad_Id_Cliente")&"'")%>
					  			(<%response.Write(RsPf.recordcount)%>) </a> 
                                	<%else%>
                                	&nbsp; 
                                	<%end if%>								</td>
								<%Set RsPfs = cConexao.Execute("Select Crs_Int_Id_Contas_Receber_Servico, CRS.Nfs_Int_Id_Nota_Fiscal_Servico, CRS.Nfs_Char_Serie_Nota_Fiscal_Servico, CRS.Cli_Int_Id_Cliente, Crs_Int_Numero_Parcela, Crs_Int_Qtd_Parcela, Crs_Dec_Valor_Receber, Crs_Dec_Juros_Dias, Crs_Date_Emissao, Crs_Date_Vencimento, Crs_Dec_Valor_Baixa, Crs_Date_Baixa, NFS.Nfs_Var_Numero from t_contas_receber_servico as CRS inner join t_nota_fiscal_SERVICO as NFS on ( CRS.Nfs_Int_Id_Nota_Fiscal_Servico = NFS.Nfs_Int_Id_Nota_Fiscal_Servico ) where CRS.Cli_Int_Id_Cliente = '"&Session("Cad_Id_Cliente")&"'")%>
                              	<td width="270" valign="middle" class="TextoFormulario"> 
                                	<a href="con_posicao_crs_cliente.asp?Opcao_Visualizacao=T"><img src="../img_sistema/botao/botao_posicao_financeira.gif" width="21" height="21" border="0"> Posição Financeira Serviço(<%=RsPfs.recordcount%>)</a>							   </td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr align="center"> 
                        <td colspan="4" class="TextoFormulario">
							<table width="800" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="198" class="TextoFormulario" valign="middle">
							  	<a href="con_servicos_mensais.asp">&nbsp;<img src="../img_sistema/botao/botao_servicos_mensais.gif" width="30" height="31" border="0">Servi&ccedil;os Mensais
							  <%Call ConexaoSistema()
								Set RsSm = cConexao.Execute("select SM_Cliente.Smc_Int_Servicos_Mensais_Cliente, SM_Cliente.Cli_Int_Id_Cliente, SM_Cliente.Psc_Int_Id_Produto_Servico, Produto_Servico.Psc_Var_Titulo, SM_Cliente.Smc_Date_Vencto, SM_Cliente.Smc_Dec_Valor, SM_Cliente.Pcs_Int_Id_Produto_Composto_Servico, Composto_Servico.Pcs_Var_Titulo, SM_Cliente.Smc_Int_Qtd from t_servicos_mensais_cliente as SM_Cliente left join t_produto_servico as Produto_Servico on ( SM_Cliente.Psc_Int_Id_Produto_Servico = Produto_servico.Psc_Int_Id_Produto_Servico ) left join t_produto_composto_servico as Composto_Servico on ( SM_Cliente.Pcs_Int_Id_Produto_Composto_Servico = Composto_Servico.Pcs_Int_Id_Produto_Composto_Servico ) where SM_Cliente.Cli_Int_Id_Cliente = '" & request.QueryString("id_cliente") & "'")%>
				  			  (<%response.Write(RsSm.recordcount)%>)							  </a></td>
								<td width="229" class="TextoFormulario" align="left" valign="middle">													                                        <%if ( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) then%>		
										<a href="con_cotacao_mercadoria_cliente.asp"> 
                                		&nbsp;<img src="../img_sistema/botao/botao_visualiza_cotacao.gif" width="21" height="21" border="0"> 
                                		Ver Cota&ccedil;&atilde;o
										<%Call ConexaoSistema() 
										Set RsCot = cConexao.Execute("select Cotacao.Ctm_Int_Id_Cot_Mercadoria, Cotacao.Ctm_Date_Data_Emissao, Pagamento.Pag_Var_Descricao, Status.Sts_Var_Descricao from t_cot_mercadoria as Cotacao inner join t_pagamento as Pagamento on ( Cotacao.Pag_Int_Id_Pagamento = Pagamento.Pag_Int_Id_Pagamento ) inner join t_status as Status on ( Cotacao.Sts_int_Id_Status = Status.Sts_int_Id_Status )  where Cotacao.Cli_Int_Id_Cliente = '"&Request.QueryString("Id_Cliente")&"' order by Cotacao.Ctm_Date_Data_Emissao Desc, Cotacao.Ctm_Int_Id_Cot_Mercadoria")%>
					  					(<%response.Write(RsCot.recordcount)%>)										</a> 
                                		<%else%>
                                		&nbsp; 
                                		<%end if%>							  </td>
								<td width="273" class="TextoFormulario" valign="middle">
									<%if ( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) then%>
                                	<a href="con_orcamento_cliente.asp">&nbsp;<img src="../img_sistema/botao/botao_visualiza_cotacao.gif" width="21" height="21" border="0"> 
                                	Ver Or&ccedil;amento Servi&ccedil;os
									<%Call ConexaoSistema()
									Set RsOs = cConexao.Execute("select Orcamento.Orc_Int_Id_Orcamento, Orcamento.Orc_Date_Emissao, Pagamento.Pag_Var_Descricao, Status.Sts_Var_Descricao from t_orcamento as Orcamento inner join t_pagamento as Pagamento on ( Orcamento.Pag_Int_Id_Pagamento = Pagamento.Pag_Int_Id_Pagamento ) inner join t_status as Status on ( Orcamento.Sts_int_Id_Status = Status.Sts_int_Id_Status ) where Orcamento.Cli_Int_Id_Cliente = '"&Request.QueryString("Id_Cliente")&"'")%>
					  				(<%response.Write(RsOs.recordcount)%>)									</a>
									<%end if%>							  </td>
							</tr>
						</table>					</td>
				</tr>
					<%else%>
						&nbsp; 
					<%end if%>
				<%end if%>
                      <%if ( Request.QueryString("Opcao") <> "2" and request.Form("id_contato") = Empty ) then%>
					  <tr>
						<td colspan="5" class="TextoFormulario"><hr width="100%" size="1"></td>
			  	      </tr>
					  <tr> 
                        <td width="133" class="TextoFormulario">Nome</td>
                        <td colspan="2" class="TextoCampoFormulario"><input name="Ctc_Var_Nome" type="text" class="TextoCampoFormulario" id="Ctc_Var_Nome" value="<%=cCtc_Var_Nome%>" size="50" maxlength="100" Opcao="texto" Campo="Nome"> <input id="contato" type="hidden" value="<%=request.QueryString("id_contato")%>" /><input id="cliente" type="hidden" value="<%=request.QueryString("id_cliente")%>" />                        </td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Cargo</td>
                        <td width="300" class="TextoCampoFormulario"> <input name="Ctc_Var_Cargo" type="text" c class="TextoCampoFormulario" id="Ctc_Var_Cargo" value="<%=cCtc_Var_Cargo%>" size="25" maxlength="100"></td>
                        <td width="304" colspan="-1" class="TextoFormulario"> 
                          Sexo 
                          <select name="Ctc_Char_Sexo"  class="TextoCampoFormulario" id="select3" opcao="select" campo="Ctc_Char_Sexo">
                            <option value="M" <%if cCtc_Char_Sexo = "M" then%>selected<%end if%>>Masculino</option>
                            <option value="F" <%if cCtc_Char_Sexo = "F" then%>selected<%end if%>>Feminino</option>
                        </select> </td>
                      </tr>
                      <tr> 
                        <td width="133" class="TextoFormulario">Vendedor</td>
                        <td class="TextoCampoFormulario"><select name="Ven_Int_Id_Vendedor2"  class="TextoCampoFormulario" id="select2" Opcao="select" Campo="Vendedor">
                            <%
							Call ConexaoSistema()
							cSql = "Select Ven_Int_Id_Vendedor, Ven_Var_Nome from t_vendedor where sts_int_id_status = '1' order by Ven_Var_Nome"
							Set RsContato = cConexao.Execute( cSql )
							if ( Not RsContato.Eof ) then
							Do While ( Not RsContato.Eof )
							%>
                            <option value="<%=RsContato("Ven_Int_Id_Vendedor")%>" <%if ( request.QueryString("opcao") = 2 and nVen_Int_id_Vendedor = RsContato("Ven_Int_Id_Vendedor") ) or (request.QueryString("opcao") = 1 and Session("Id_Vendedor") = RsContato("Ven_Int_Id_Vendedor")) then%>selected<%end if%>><%=RsContato("Ven_Var_Nome")%></option>
                            <%
							RsContato.MoveNext
							Loop
							end if
							RsContato.Close : Set RsContato = nothing : Call FechaConexaoSistema()
							%>
                          </select> </td>
                        <td colspan="-1" class="TextoFormulario">Status 
                          <select name="Sgc_Int_Id_Status"  class="TextoCampoFormulario" id="select7" Opcao="select" Campo="Status Cliente">
                            <%
							Call ConexaoSistema()
							cSql = "Select Sts_Int_Id_Status, Sts_Var_Descricao, Sts_Char_Contato from t_status "
							cSql = cSql & "where Sts_Char_Contato = " & "'T'" & "order by Sts_int_id_status desc"
							Set RsContato = cConexao.Execute( cSql )
							if ( Not RsContato.Eof ) then
							Do While ( Not RsContato.Eof )
							%>
							<%if RsContato("Sts_Int_Id_Status") <> 0 then %>
                            <option value="<%=RsContato("Sts_Int_Id_Status")%>"							
								<%if ( nSts_Int_Id_Status = RsContato("Sts_Int_Id_Status") ) then%>
								selected="nSts_Int_Id_Status"
								<%end if%>><%=RsContato("Sts_Var_Descricao")%>							</option>
                            <%
							end if
							RsContato.MoveNext
							Loop
							end if
							RsContato.Close : Set RsContato = Nothing : Call FechaConexaoSistema()
							%>
                          </select></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Data de Nascimento</td>
                        <td class="TextoCampoFormulario"><input name="Ctc_Int_Data_Nascimento" type="text" c class="TextoCampoFormulario" id="Ctc_Int_Data_Nascimento" value="<%=dCtc_Int_Data_Nascimento%>" size="15" maxlength="12">
                          (dd/mm/aaaa)</td>
                        <td colspan="-1" class="TextoCampoFormulario">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">DDD </td>
                        <td colspan="2" class="TextoFormulario"><table width="595" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="28"><input name="Ctc_Char_Ddd" type="text" class="TextoCampoFormulario" id="Ctc_Char_Ddd" value="<%=cCtc_Char_Ddd%>" size="4" maxlength="3" Opcao="texto" Campo="DDD"></td>
                              <td width="62" class="TextoFormulario">Telefone                              </td>
                              <td width="104" class="TextoFormulario"><input name="Ctc_Var_Telefone" type="text" class="TextoCampoFormulario" id="Ctc_Var_Telefone" value="<%=cCtc_Var_Telefone%>" size="12" maxlength="9" Opcao="texto" Campo="Telefone" onKeyPress="return txtBoxFormat(this, '9999-9999', event);"></td>
                              <td width="96" class="TextoFormulario">Ramal 
                                <input name="Ctc_Var_Ramal" type="text" class="TextoCampoFormulario" id="Ctc_Var_Ramal" value="<%=cCtc_Var_Ramal%>" size="4" maxlength="5">                              </td>
                              <td width="136" class="TextoFormulario">Fax 
                                <input name="Ctc_Var_Fax" type="text" class="TextoCampoFormulario" id="Ctc_Var_Fax" value="<%=cCtc_Var_Fax%>" size="12" maxlength="9" onKeyPress="return txtBoxFormat(this, '9999-9999', event);">                              </td>
                              <td width="169" class="TextoFormulario">Ramal 
                                <input name="Ctc_Var_Ramal_Fax" type="text" class="TextoCampoFormulario" id="Ctc_Var_Ramal_Fax" value="<%=cCtc_Var_Ramal_Fax%>" size="4" maxlength="5">                              </td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">DDD Celular</td>
                        <td colspan="2" class="TextoFormulario"><table width="194" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="28"><input name="Ctc_Char_Ddd_Celular" type="text" class="TextoCampoFormulario" id="Ctc_Char_Ddd_Celular" value="<%=cCtc_Char_Ddd_Celular%>" size="4" maxlength="3" opcao="texto" campo="DDD"></td>
                              <td width="62" class="TextoFormulario">Celular</td>
                              <td width="104" class="TextoFormulario"><input name="Ctc_Var_Celular" type="text" c class="TextoCampoFormulario" id="Ctc_Var_Celular2" value="<%=cCtc_Var_Celular%>" size="12" maxlength="9" onKeyPress="return txtBoxFormat(this, '9999-9999', event);"></td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">E-mail</td>
                        <td colspan="2" class="TextoFormulario"> <input name="Ctc_Var_email" type="text" c class="TextoCampoFormulario" id="Ctc_Var_email" value="<%=cCtc_Var_email%>" size="50" maxlength="256" Opcao="email" Campo="E-mail"></td>
                      </tr>
                      <tr> 
                        <td class="TextoFormulario">Contato Internet</td>
                        <td colspan="2"><input name="Ctc_Char_Flag_Internet" type="checkbox" class="TextoCampoFormulario" id="Ctc_Char_Flag_Internet" value="T" <%if ( lCtc_Char_Flag_Internet ) then%>checked<%end if%>></td>
                      </tr>
					  <tr> 
                        <td class="TextoFormulario">Mostra Part Number Site</td>
                        <td colspan="2" class="TextoFormulario"><input name="Ctc_Char_Flag_PN" type="checkbox" class="TextoCampoFormulario" id="Ctc_Char_Flag_PN" value="T" <%if ( cCtc_Char_Flag_PN ) then%>checked<%end if%>></td>
                      </tr>
					  <tr>
						<td colspan="5" class="TextoFormulario"><hr width="100%" size="1"></td>
			  	      </tr>
                      <%if ( VerificaPermissaoBD( Session("Id_Login"), 1, "CLI01" ) ) then%>
                      <tr align="center"> 
                        <td colspan="4"><input name="imageField" type="image" src="../img_sistema/botao/botao_inclusao.gif" width="34" height="32" border="0">                        </td>
                      </tr>
                      <%end if%>
                      <%else%>
                      <%if ((nSts_Int_Id_Status = 6 and Session("Id_Login_Master")) or (nSts_Int_Id_Status <> 6)) and (( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) )or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) ) then%>
                      <%end if%>
                      <%if ( VerificaPermissaoBD( Session("Id_Login"), 2, "CLI01" ) or ( Vvendedor = 0 ) ) then%>
                      <tr align="center"> 
                        <td colspan="4">
						<%if ((nSts_Int_Id_Status = 6 and Session("Id_Login_Master")) or (nSts_Int_Id_Status <> 6)) and (( Session("Id_Vendedor") = nVen_Int_id_Vendedor ) or ( ValidaUsuario( nVen_Int_id_Vendedor, Session("Id_Vendedor") ) or ( VerificaPermissaoBD( Session("Id_Login"), 4, "CLI07" ) ) or ( Vvendedor = 0 ) or VerificaPermissaoAvulsaBD(request.QueryString("Id_Cliente"), Session("Id_Vendedor")) ) ) then%><input name="imageField" type="image" src="../img_sistema/botao/botao_alteracao.gif" width="34" height="32" border="0"><%end if%></td>
                      </tr>
                      <%end if%>
                      <%end if%>
                    </table>
                  </form></td>
              </tr>
            </table>
			
          </td>
        </tr>
        <tr>
		  	<!--#Include File="../net/footer.net"-->
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<% 

if ( Request.QueryString("Opcao") ) = 2 then%>
<%Call ConexaoSistema()%>
<%if Verifica_Integridade_Cliente(Session("Cad_Id_Cliente"),false) = "EC" then%>
<script language="JavaScript">
	alert("Cadastro de cliente incompleto !!!\n" + "Cadastre um ENDEREÇO e um CONTATO !!!");
</script>
<%elseif Verifica_Integridade_Cliente(Session("Cad_Id_Cliente"),false) = "E" then%>
<script language="JavaScript">
	alert("Cadastro de cliente incompleto !!!\n" + "Cadastre um ENDEREÇO !!!");
</script>
<%elseif Verifica_Integridade_Cliente(Session("Cad_Id_Cliente"),false) = "C" then%>
<script language="JavaScript">
	alert("Cadastro de cliente incompleto !!!\n" + "Cadastre um CONTATO !!!");
</script>
<%end if%>

<% 
Call FechaConexaoSistema()%>
<%end if%>
