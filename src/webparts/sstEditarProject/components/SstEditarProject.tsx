import * as React from 'react';
import styles from './SstEditarProject.module.scss';
import { ISstEditarProjectProps } from './ISstEditarProjectProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jquery from 'jquery';
import * as $ from "jquery";
import * as jQuery from "jquery";
import { sp, IItemAddResult, DateTimeFieldFormatType } from "@pnp/sp/presets/all";
import "bootstrap";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { Web } from "sp-pnp-js";
import pnp from "sp-pnp-js";
import { ICamlQuery } from '@pnp/sp/lists';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { allowOverscrollOnElement, DatePicker } from 'office-ui-fabric-react';
import { UrlQueryParameterCollection, Version } from '@microsoft/sp-core-library';
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { SiteUser } from 'sp-pnp-js/lib/sharepoint/siteusers';
import BootstrapTable from 'react-bootstrap-table-next';
import filterFactory, { textFilter } from 'react-bootstrap-table2-filter';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _web;
var _caminho;
var _arrOwner = [];
var _arrOwnerID = [];
var _arrParticipants = [];
var _arrParticipantsID = [];
var _productDescription;
var _criticalRequirements;
var _projectID;
var _arrAprovadorEngenharia = [];
var _arrCliente = [];
var _pos = 0;
var _pos2 = 0;

export interface IReactGetItemsState {

  itemsCliente: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsListAnexosItem: [
    {
      "FileName": any,
      "ServerRelativeUrl": any,
    }
  ],
  itemsListAnexos: [
    {
      "Name": any,
      "ServerRelativeUrl": any,
    }
  ],
  itemsCategoria: [],
  itemsTipo: [],
  valorItemsCategoria: "",
  valorItemsTipo: "",
  PeoplePickerDefaultItemsOwner: string[],
  PeoplePickerDefaultItemsParticipants: string[],
  itemsListRelatedMilestones: any[],
  itemsListRelatedIssues: any[],

}


const tablecolumnsRelatedIssues = [
  {
    dataField: "Title",
    text: "Title",
    headerStyle: { "backgroundColor": "#bee5eb" },
    classes: 'headerPreStage',
    headerClasses: 'text-center',
  },

]

const tablecolumnsRelatedMilestones = [
  {
    dataField: "Title",
    text: "Title",
    headerStyle: { "backgroundColor": "#bee5eb" },
    classes: 'headerPreStage',
    headerClasses: 'text-center',
  },

]


export default class SstEditarProject extends React.Component<ISstEditarProjectProps, IReactGetItemsState> {

  public constructor(props: ISstEditarProjectProps, state: IReactGetItemsState) {
    super(props);
    this.state = {

      itemsCliente: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsListAnexosItem: [
        {
          "FileName": "",
          "ServerRelativeUrl": ""
        }
      ],
      itemsListAnexos: [
        {
          "Name": "",
          "ServerRelativeUrl": "",
        }
      ],
      itemsCategoria: [],
      itemsTipo: [],
      valorItemsCategoria: "",
      valorItemsTipo: "",
      PeoplePickerDefaultItemsOwner: [],
      PeoplePickerDefaultItemsParticipants: [],
      itemsListRelatedMilestones: [],
      itemsListRelatedIssues: [],
    };
  }


  public componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);
    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _projectID = parseInt(queryParms.getValue("ProjectID"));


    jQuery("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/SiteAssets/loading.gif'/>
      <br/>Aguarde....<br/><br/>
      Dependendo do tamanho do anexo e a velocidade<br>
       da Internet essa ação pode demorar um pouco. <br>
       Não fechar a janela!<br/><br/>`);

    this.handler();
    this.getProject();
    this.getDefaultUsers();
    this.getAnexos();


  }


  public render(): React.ReactElement<ISstEditarProjectProps> {
    return (

      <><div id="container">
        <div id="accordion">

          <div className="card">
            <div className="card-header btn" id="headingProjectInformation" data-toggle="collapse" data-target="#collapseProjectInformation" aria-expanded="true" aria-controls="collapseProjectInformation">
              <h5 className="mb-0 text-info">
                Informações da Solicitação
              </h5>
            </div>
            <div id="collapseProjectInformation" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md-6">
                      <label htmlFor="txtName">Nome</label><span className="required"> *</span>
                      <input type="text" className="form-control" id="txtName" />
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="ddlCategory">Categoria</label><span className="required"> *</span>
                      <select id="ddlCategory" className="form-control" value={this.state.valorItemsCategoria} onChange={(e) => this.onChangeCategoria(e.target.value)}>
                        <option value="0" selected>Selecione...</option>
                        {this.state.itemsCategoria.map(function (item, key) {
                          return (
                            <option value={item}>{item}</option>
                          );
                        })}
                      </select>
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="ddlType">Tipo</label><span className="required"> *</span>
                      <select id="ddlType" className="form-control" value={this.state.valorItemsTipo} onChange={(e) => this.onChangeTipo(e.target.value)}>
                        <option value="0" selected>Selecione...</option>
                        {this.state.itemsTipo.map(function (item, key) {
                          return (
                            <option value={item}>{item}</option>
                          );
                        })}
                      </select>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtOwner">Owner</label><span className="required"> *</span>
                      <PeoplePicker
                        context={this.props.context as any}
                        //titleText="Aprovador Engenharia"
                        personSelectionLimit={1}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        onChange={this._getPeoplePickerOwner.bind(this)}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        defaultSelectedUsers={this.state.PeoplePickerDefaultItemsOwner}
                        ensureUser={true} />

                    </div>
                    <div className="form-group col-md">
                      <label htmlFor="txtParticipants">Participantes</label><span className="required"> *</span>
                      <PeoplePicker
                        context={this.props.context as any}
                        //titleText="Aprovador Engenharia"
                        personSelectionLimit={20}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        onChange={this._getPeoplePickerParticipants.bind(this)}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        defaultSelectedUsers={this.state.PeoplePickerDefaultItemsParticipants}
                        ensureUser={true} />
                    </div>
                  </div>
                </div>


                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtProduct">Descrição do Produto / Serviço</label><span className="required"> *</span>
                      <div id='richTextDescricaoProduto'>
                        <RichText className="editorRichTex" value=""
                          onChange={(text) => this.onTextChangeProductDescription(text)} />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtCritical">Requisitos críticos</label><span className="required"> *</span>
                      <div id='richTextRequisitosCriticos'>
                        <RichText className="editorRichTex" value=""
                          onChange={(text) => this.onTextChangeCriticalRequirements(text)} />
                      </div>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="ddlProduto">Cliente</label><span className="required"> *</span>
                      <table>
                        <tr>
                          <td>
                            <div className="col-sm-6">
                              <select multiple={true} id='ddlCliente1' className="form-control" name="ddlCliente1" style={{ "height": "194px", "width": "350px" }}>

                                {this.state.itemsCliente.map(function (item, key) {

                                  if (_arrCliente.indexOf(item.ID) == -1) {
                                    return (
                                      <option className="optCliente" value={item.ID}>{item.Title}</option>
                                    );
                                  }

                                })}

                              </select>
                            </div>
                          </td>
                          <td>
                            <div>
                              <input type="button" onClick={this.addButtonArea} className="btn btn-light" id="addButtonArea" value="Adicionar >" alt="Salvar" /></div><br />
                            <input type="button" onClick={this.removeButtonArea} className="btn btn-light" id="removeButtonArea" value="< Remover"
                              alt="Salvar" />
                          </td>
                          <td>
                            <div className="col-sm-6">
                              <select multiple={true} id="ddlCliente2" className="form-control" name="ddlCliente2" style={{ "height": "194px", "width": "350px" }}>
                                {this.state.itemsCliente.map(function (item, key) {

                                  if (_arrCliente.indexOf(item.ID) !== -1) {
                                    return (
                                      <option className="optCliente" value={item.ID}>{item.Title}</option>
                                    );
                                  }

                                })}
                              </select>
                            </div>
                          </td>
                        </tr>
                      </table>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtOMPDocuments">Documentos OMP</label><span className="required"> *</span>
                      <input type="text" className="form-control" id="txtOMPDocuments" />
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </div>


          <div className="card">
            <div className="card-header btn" id="headingInformacoesProduto" data-toggle="collapse" data-target="#collapseInformacoesProduto" aria-expanded="true" aria-controls="collapseInformacoesProduto">
              <h5 className="mb-0 text-info">
                Anexos
              </h5>
            </div>
            <div id="collapseInformacoesProduto" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row ">
                    <div className="form-group col-md" >
                      <label htmlFor="txtTitulo">Anexo </label><br></br>
                      <input className="multi" data-maxsize="1024" type="file" id="input" multiple />
                    </div>
                    <div className="form-group col-md" >

                    </div>

                  </div>
                  <br />
                  <p className='text-info'>Total máximo permitido: 15 MB</p>

                </div>

                <div className="form-group">
                  <div className="form-row ">
                    <div className="form-group col-md" >
                      {this.state.itemsListAnexosItem.map((item, key) => {

                        _pos++;
                        var txtAnexoItem = "anexoItem" + _pos;
                        var btnExcluirAnexoitem = "btnExcluirAnexoitem" + _pos;

                        var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Anexos')/items('${_projectID}')/AttachmentFiles`;
                        url = this.props.siteurl;

                        var caminho = `${url}/Lists/Documentos/Attachments/${_projectID}/${item.FileName}`;

                        return (

                          <><a id={txtAnexoItem} target='_blank' data-interception="off" href={caminho} title="">{item.FileName}</a><a style={{ "cursor": "pointer" }} onClick={() => this.excluirAnexoItem(`${item.ServerRelativeUrl}`, `${item.FileName}`, `${txtAnexoItem}`, `${btnExcluirAnexoitem}`)} id={btnExcluirAnexoitem}>&nbsp;Excluir</a><br></br></>


                        );



                      })}
                      {this.state.itemsListAnexos.map((item, key) => {

                        _pos2++;
                        //var txtAnexoItem = "anexoItem" + _pos;
                        //var btnExcluirAnexoitem = "btnExcluirAnexoitem" + _pos;

                        //var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Documentos')/items('${_idOMP}')/AttachmentFiles`;
                        //url = this.props.siteurl;

                        var caminho = item.ServerRelativeUrl;

                        var btnExcluirAnexo2 = `btnExcluirAnexo2${_pos2}`;
                        var txtAnexoItem2 = `anexo2${_pos2}`;

                        var relativeURL = window.location.pathname;
                        var url = window.location.pathname;
                        var nomePagina = url.substring(url.lastIndexOf('/') + 1);
                        var strRelativeURL = relativeURL.replace(`SitePages/${nomePagina}`, "");

                        return (

                          <><a id={txtAnexoItem2} target='_blank' data-interception="off" href={caminho} title="">{item.Name}</a><a style={{ "cursor": "pointer" }} onClick={() => this.excluirAnexo(`${strRelativeURL}/Anexos/${_projectID}`, `${item.Name}`, `${txtAnexoItem2}`, `${btnExcluirAnexo2}`)} id={btnExcluirAnexo2}>&nbsp;Excluir</a><br></br></>

                        );



                      })}
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingPreStageHardware" data-toggle="collapse" data-target="#collapsePreStageHardware" aria-expanded="true" aria-controls="collapsePreStageHardware">
              <h5 className="mb-0 text-info">
                Related Issues
              </h5>
            </div>
            <div id="collapsePreStageHardware" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">
                <div id='tabelaPreStageSoftware'>
                  <BootstrapTable bootstrap4 striped responsive condensed hover={false} className="gridTodosItens" id="gridTodosItensRelatedIssues" keyField='id' data={this.state.itemsListRelatedIssues} columns={tablecolumnsRelatedIssues} headerClasses="header-class" />
                </div>
                <button id='btnAbrirModalCadastrarPreStage' className="btn btn-secondary btnCustom btn-sm">Adicionar</button>&nbsp;
                <button id='btnAbrirModalCadastrarPreStageEmLote' className="btn btn-secondary btnCustom btn-sm">Adicionar em lote</button>
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingPreStageHardware" data-toggle="collapse" data-target="#collapsePreStageHardware" aria-expanded="true" aria-controls="collapsePreStageHardware">
              <h5 className="mb-0 text-info">
              Related Milestones
              </h5>
            </div>
            <div id="collapsePreStageHardware" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">
                <div id='tabelaPreStageSoftware'>
                  <BootstrapTable bootstrap4 striped responsive condensed hover={false} className="gridTodosItens" id="gridTodosItensPreStageSoftware" keyField='id' data={this.state.itemsListRelatedMilestones} columns={tablecolumnsRelatedMilestones} headerClasses="header-class" />
                </div>
                <button id='btnAbrirModalCadastrarPreStage' className="btn btn-secondary btnCustom btn-sm">Adicionar</button>&nbsp;
                <button id='btnAbrirModalCadastrarPreStageEmLote' className="btn btn-secondary btnCustom btn-sm">Adicionar em lote</button>
              </div>
            </div>
          </div>




        </div>
      </div>

        <br></br><div className="text-right">
          <button id="btnConfirmarSalvar" className="btn btn-success">Editar</button>
        </div>





      </>


    );
  }


  protected async handler() {

    var reactCategoria = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('Projects List')/fields?$filter=EntityPropertyName eq 'ProjCategory'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactCategoria.setState({
          itemsCategoria: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    var reactTipo = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('Projects List')/fields?$filter=EntityPropertyName eq 'Project_x0020_type'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactTipo.setState({
          itemsTipo: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    var reactHandlerClientes = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Client')/items?$top=4999&$&$orderby= Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactHandlerClientes.setState({
          itemsCliente: resultData.d.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });



  }


  protected getProject() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Projects List')/items?$select=ID,Title,ProjCategory,Project_x0020_type,AssignedTo/ID,AssignedTo/Title,Participants/ID,Participants/Title,Product_x0020_description_x0020_,Critical_x0020_requirements,Client/ID,OMP_x0020_documents&$expand=AssignedTo,Participants,Client&$filter=ID eq ` + _projectID,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        //  console.log("resultData doc", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var nome = resultData.d.results[i].Title;
            var category = resultData.d.results[i].ProjCategory;
            var tipo = resultData.d.results[i].Project_x0020_type;
            var omp = resultData.d.results[i].OMP_x0020_documents;

            if (resultData.d.results[i].AssignedTo.hasOwnProperty('results')) {

              for (var x = 0; x < resultData.d.results[i].AssignedTo.results.length; x++) {

                _arrOwner.push(resultData.d.results[i].AssignedTo.results[x].Title);
                _arrOwnerID.push(resultData.d.results[i].AssignedTo.results[x].ID);

              }

            }

            if (resultData.d.results[i].Participants.hasOwnProperty('results')) {

              for (var x = 0; x < resultData.d.results[i].Participants.results.length; x++) {

                _arrParticipants.push(resultData.d.results[i].Participants.results[x].Title);
                _arrParticipantsID.push(resultData.d.results[i].Participants.results[x].ID);


              }

            }

            jQuery("#txtName").val(nome);
            jQuery("#txtOMPDocuments").val(omp);

            this.setState({
              valorItemsCategoria: category,
              valorItemsTipo: tipo

            });


            var descricaoProduto = resultData.d.results[i].Product_x0020_description_x0020_;
            var txtDescricaoProduto = "";
            if (descricaoProduto != null) {
              txtDescricaoProduto = descricaoProduto.replace(/<[\/]{0,1}(div)[^><]*>/g, "");
              console.log("txtParametros", txtDescricaoProduto);
              if (txtDescricaoProduto.includes("<font")) {
                txtDescricaoProduto = txtDescricaoProduto.replace("font", "span");
                txtDescricaoProduto = txtDescricaoProduto.replace("font", "span");
              }
              if (txtDescricaoProduto.includes("color")) {
                txtDescricaoProduto = txtDescricaoProduto.replace('color="', 'style="color:');
              }
              txtDescricaoProduto = txtDescricaoProduto.trim();
            }
            jQuery('#richTextDescricaoProduto').find('.ql-editor').html(`${txtDescricaoProduto}`);


            var requisitosCriticos = resultData.d.results[i].Critical_x0020_requirements;
            var txtRequisitosCriticos = "";
            if (requisitosCriticos != null) {
              txtRequisitosCriticos = requisitosCriticos.replace(/<[\/]{0,1}(div)[^><]*>/g, "");
              console.log("txtParametros", txtRequisitosCriticos);
              if (txtRequisitosCriticos.includes("<font")) {
                txtRequisitosCriticos = txtRequisitosCriticos.replace("font", "span");
                txtRequisitosCriticos = txtRequisitosCriticos.replace("font", "span");
              }
              if (txtRequisitosCriticos.includes("color")) {
                txtRequisitosCriticos = txtRequisitosCriticos.replace('color="', 'style="color:');
              }
              txtRequisitosCriticos = txtRequisitosCriticos.trim();
            }
            jQuery('#richTextRequisitosCriticos').find('.ql-editor').html(`${txtRequisitosCriticos}`);


            var arrCliente = [];
            arrCliente = resultData.d.results[i].Client.results;
            var tamArrCliente = resultData.d.results[i].Client.results.length;

            for (i = 0; i < tamArrCliente; i++) {

              _arrCliente.push(arrCliente[i].ID);

            }


          }
        }




      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }

    })

  }

  protected async getAnexos() {

    var montaImagem = "";
    var montaOutros = "";

    var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Projects List')/items('${_projectID}')/AttachmentFiles`;
    var _url = this.props.siteurl;
    // console.log("url", url);

    $.ajax
      ({
        url: url,
        method: "GET",
        async: false,
        headers:
        {
          // Accept header: Specifies the format for response data from the server.
          "Accept": "application/json;odata=verbose"
        },
        success: async (resultData) => {

          var dataresults = resultData.d.results;

          var reactHandler = this;

          reactHandler.setState({
            itemsListAnexosItem: dataresults
          });

        },
        error: function (xhr, status, error) {
          console.log("Falha anexo");
        }
      }).catch((error: any) => {
        console.log("Erro Anexo do item: ", error);
      });


    var relativeURL = window.location.pathname;
    var url = window.location.pathname;
    var nomePagina = url.substring(url.lastIndexOf('/') + 1);
    var strRelativeURL = relativeURL.replace(`SitePages/${nomePagina}`, "");

    await _web.getFolderByServerRelativeUrl(`${strRelativeURL}/Anexos/${_projectID}`).files.orderBy('TimeLastModified', true)

      .expand('ListItemAllFields', 'Author').get().then(r => {

        console.log("r", r);

        var reactHandler = this;

        reactHandler.setState({
          itemsListAnexos: r
        });

      }).catch((error: any) => {
        console.log("Erro onChangeCliente: ", error);
      });


  }



  async excluirAnexoItem(ServerRelativeUr, name, elemento, elemento2) {

    if (confirm("Deseja realmente excluir o arquivo " + name + "?") == true) {

      var relativeURL = window.location.pathname;
      var url = window.location.pathname;
      var nomePagina = url.substring(url.lastIndexOf('/') + 1);
      var strRelativeURL = relativeURL.replace(`SitePages/${nomePagina}`, "");

      await _web.getFolderByServerRelativeUrl(`${strRelativeURL}/Lists/Documentos/Attachments/${_projectID}`).files.getByName(name).delete()
        .then(async response => {
          jQuery(`#${elemento}`).hide();
          jQuery(`#${elemento2}`).hide();
          alert("Arquivo excluido com sucesso.");
        }).catch(console.error());

    } else {
      return false;
    }
  }


  async excluirAnexo(ServerRelativeUr, name, elemento, elemento2) {


    if (confirm("Deseja realmente excluir o arquivo " + name + "?") == true) {

      //  console.log("ServerRelativeUr", ServerRelativeUr);
      //  console.log("name", name);
      await _web.getFolderByServerRelativeUrl(ServerRelativeUr).files.getByName(name).delete()
        .then(async response => {
          jQuery(`#${elemento}`).hide();
          jQuery(`#${elemento2}`).hide();
          alert("Arquivo excluido com sucesso.");
        }).catch(console.error());

    } else {
      return false;
    }

  }

  private getDefaultUsers() {

    setTimeout(() => {

      console.log("_arrOwner", _arrOwner);

      this.setState({
        PeoplePickerDefaultItemsOwner: _arrOwner,
        PeoplePickerDefaultItemsParticipants: _arrParticipants,

      });

    }, 2000);


  }


  private _getPeoplePickerOwner(items: any[]) {
    _arrOwner = items;
    console.log('Items:', items);
  }

  private _getPeoplePickerParticipants(items: any[]) {
    _arrParticipants = items;
    console.log('Items:', items);
  }

  private onTextChangeProductDescription = (newText: string) => {
    _productDescription = newText;
    return newText;
  }

  private onTextChangeCriticalRequirements = (newText: string) => {
    _criticalRequirements = newText;
    return newText;
  }

  protected addButtonArea = () => {
    var $options = $('#ddlCliente1 option:selected');
    $options.appendTo("#ddlCliente2");
  }

  protected removeButtonArea = () => {
    var $options = $('#ddlCliente2 option:selected');
    $options.appendTo("#ddlCliente1");
  }

  private onChangeCategoria = (val) => {
    this.setState({
      valorItemsCategoria: val,
    });
  }

  private onChangeTipo = (val) => {
    this.setState({
      valorItemsTipo: val,
    });
  }





}
