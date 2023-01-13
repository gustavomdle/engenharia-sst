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
import InputMask from 'react-input-mask';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _web;
var _caminho;
var _arrOwner = [];
var _arrOwnerID = [];
var _arrParticipants = [];
var _arrParticipantsID = [];
var _productDescription;
var _descricaoRelatedIssues;
var _descricaoComentariosRelatedIssues;
var _criticalRequirements;
var _projectID;
var _arrAprovadorEngenharia = [];
var _arrCliente = [];
var _pos = 0;
var _pos2 = 0;
var _atribuidoARelatedIssues = [];
var _descricaoComentariosMilestone;
var _idRelatedIssues;
var _arrValAtribuidAID = [];
var _atribuidoARelatedIssuesTeveAlteracao = false;
var _ownerTevelateracao = false;
var _participantsTevelateracao = false;
var _idRelatedMilestones;
var _size: number = 0;
var _pastaCriada = "";
var _grupos;
var _statusAtual;
var _siteNovo;
var _projectTitle;

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
  itemsIssueStatus: [],
  itemsPriority: [],
  itemsIssueCategoria: [],
  valorItemsCategoria: "",
  valorItemsTipo: "",
  PeoplePickerDefaultItemsOwner: string[],
  PeoplePickerDefaultItemsParticipants: string[],
  itemsListRelatedMilestones: any[],
  itemsListRelatedIssues: any[],
  PeoplePickerAtribudioARelatedIssues: string[];
  PeoplePickerAtribudioARelatedIssuesEditar: string[];
  valorItemsDataVencimentoRelatedIssues: any,
  valorItemsDataVencimentoRelatedMilestones: any,
  valorProjectMilestoneRelatedMilestones: any,

}




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
      itemsIssueStatus: [],
      itemsPriority: [],
      itemsIssueCategoria: [],
      valorItemsCategoria: "",
      valorItemsTipo: "",
      PeoplePickerDefaultItemsOwner: [],
      PeoplePickerDefaultItemsParticipants: [],
      itemsListRelatedMilestones: [],
      itemsListRelatedIssues: [],
      PeoplePickerAtribudioARelatedIssues: [],
      PeoplePickerAtribudioARelatedIssuesEditar: [],
      valorItemsDataVencimentoRelatedIssues: "",
      valorItemsDataVencimentoRelatedMilestones: "",
      valorProjectMilestoneRelatedMilestones: "",
    };
  }


  public async componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);
    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    var queryParms = new UrlQueryParameterCollection(window.location.href);
    _projectID = parseInt(queryParms.getValue("ProjectID"));

    document
      .getElementById("btnAbrirModalCadastrarRelatedIssues")
      .addEventListener("click", (e: Event) => this.abrirModalCadastrarRelatedIssues());

    document
      .getElementById("btnAbrirModalCadastrarMilestones")
      .addEventListener("click", (e: Event) => this.abrirModalCadastrarMilestones());

    document
      .getElementById("btnCadastrarRelatedIssues")
      .addEventListener("click", (e: Event) => this.cadastrarProjectIssues());

    document
      .getElementById("btnEditarRelatedIssues")
      .addEventListener("click", (e: Event) => this.editarProjectIssues());

    document
      .getElementById("btnCadastrarMilestone")
      .addEventListener("click", (e: Event) => this.cadastrarMilestone());

    document
      .getElementById("btnEditarMilestone")
      .addEventListener("click", (e: Event) => this.editarMilestone());

    document
      .getElementById("btnSucessoCadastrarRelatedIssues")
      .addEventListener("click", (e: Event) => this.fecharSucessoRelatedIssues());

    document
      .getElementById("btnSucessoEditarRelatedIssues")
      .addEventListener("click", (e: Event) => this.fecharSucessoRelatedIssues());

    document
      .getElementById("btnSucessoExcluirRelatedIssues")
      .addEventListener("click", (e: Event) => this.fecharSucessoRelatedIssues());

    document
      .getElementById("btnSucessoExcluirMilestone")
      .addEventListener("click", (e: Event) => this.fecharSucessoRelatedMilestone());

    document
      .getElementById("btnSucessoCadastrarMilestone")
      .addEventListener("click", (e: Event) => this.fecharSucessoRelatedMilestone());

    document
      .getElementById("btnSucessoEditarMilestone")
      .addEventListener("click", (e: Event) => this.fecharSucessoRelatedMilestone());

    document
      .getElementById("btnConfirmarSalvar")
      .addEventListener("click", (e: Event) => this.validar("Salvar"));

    document
      .getElementById("btnConfirmarFluxo")
      .addEventListener("click", (e: Event) => this.validar("Aprovar"));

    document
      .getElementById("btnSalvar")
      .addEventListener("click", (e: Event) => this.editar("Salvar"));

    document
      .getElementById("btnIniciarFluxo")
      .addEventListener("click", (e: Event) => this.editar("Aprovar"));

    document
      .getElementById("btnSucessoSalvar")
      .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

      document
      .getElementById("btnSucessoAprovar")
      .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Aprovar"));
      
    document
      .getElementById("btnVoltar")
      .addEventListener("click", (e: Event) => this.voltar());


    jQuery("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/SiteAssets/loading.gif'/>
      <br/>Aguarde....<br/><br/>
      Dependendo do tamanho do anexo e a velocidade<br>
       da Internet essa ação pode demorar um pouco. <br>
       Não fechar a janela!<br/><br/>`);

    await _web.currentUser.get().then(f => {
      console.log("user", f);
      var id = f.Id;

      var grupos = [];

      jQuery.ajax({
        url: `${this.props.siteurl}/_api/web/GetUserById(${id})/Groups`,
        type: "GET",
        headers: { 'Accept': 'application/json; odata=verbose;' },
        async: false,
        success: async function (resultData) {

          console.log("resultDataGrupo", resultData);

          if (resultData.d.results.length > 0) {

            for (var i = 0; i < resultData.d.results.length; i++) {

              grupos.push(resultData.d.results[i].Title);

            }

          }

        },
        error: function (jqXHR, textStatus, errorThrown) {
          console.log(textStatus);
        }

      })

      console.log("grupos", grupos);
      _grupos = grupos;

    })

    jQuery("#btnConfirmarSalvar").hide();
    jQuery("#btnAbrirModalCadastrarRelatedIssues").hide();
    jQuery("#btnAbrirModalCadastrarMilestones").hide();
    jQuery("#btnConfirmarFluxo").hide();

    this.getProject();
    this.getDefaultUsers();
    this.getAnexos();
    this.handler();


  }




  public render(): React.ReactElement<ISstEditarProjectProps> {

    const tablecolumnsRelatedIssues = [
      {
        dataField: "Title",
        text: "Título",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Priority",
        text: "Prioridade",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Status",
        text: "Status",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Category",
        text: "Categoria",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "",
        text: "Atribuído a",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var atribuidoA;
          if (_siteNovo) {
            atribuidoA = row.AssignedTo.results[0].Title;
          }
          else {
            atribuidoA = row.Assigned_x0020_To_x0020_2;
          }
          console.log("atribuidoA", atribuidoA);
          return atribuidoA;
        }
      },
      {
        dataField: "DueDate",
        text: "Vencimento",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var data = new Date(row.DueDate);
          var dtdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
          return dtdata;
        }
      },
      {
        dataField: "Comment",
        text: "Descrição",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          return <div dangerouslySetInnerHTML={{ __html: `${row.Comment}` }} />;

        }
      },
      {
        dataField: "V3Comments",
        text: "Comentários",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var comentarios = row.V3Comments;
          var vlrComentario = "";

          if (comentarios != null) vlrComentario = row.V3Comments;

          return <div dangerouslySetInnerHTML={{ __html: `${vlrComentario}` }} />;

        }
      },
      {
        dataField: "",
        text: "",
        headerStyle: { "backgroundColor": "#bee5eb", "width": "130px" },
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          console.log("row related issues", row);
          var id = row.ID;
          var titulo = row.Title;
          var status = row.Status;
          var arrAtribuidA = row.AssignedTo;
          var prioridade = row.Priority;
          var descricao = row.Comment;
          var categoria = row.Category;
          var arrValAtribuidA = [];
          var arrValAtribuidAID = [];
          var comentarios = row.V3Comments;

          var data = new Date(row.DueDate);
          var formdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();

          var txtDescricao = "";

          if (descricao != null) {

            txtDescricao = descricao.replace(/<[\/]{0,1}(div)[^><]*>/g, "");
            // console.log("txtParametros", txtDescricao);

            if (txtDescricao.includes("<font")) {

              txtDescricao = txtDescricao.replace("font", "span");
              txtDescricao = txtDescricao.replace("font", "span");

            }

            if (txtDescricao.includes("color")) {

              txtDescricao = txtDescricao.replace('color="', 'style="color:');

            }

            txtDescricao = txtDescricao.trim();

            console.log("txtDescricao", txtDescricao);

          }

          var txtComentarios = "";

          if (comentarios != null) {

            txtComentarios = comentarios.replace(/<[\/]{0,1}(div)[^><]*>/g, "");

            if (txtComentarios.includes("<font")) {

              txtComentarios = txtComentarios.replace("font", "span");
              txtComentarios = txtComentarios.replace("font", "span");

            }

            if (txtComentarios.includes("color")) {

              txtComentarios = txtComentarios.replace('color="', 'style="color:');

            }

            txtComentarios = txtComentarios.trim();

            console.log("txtComentarios", txtComentarios);

          }

          console.log("formdata", formdata);

          return (

            <div>
              <button onClick={async () => { this.excluirRelatedIssues(id, titulo) }} className="btn btn-info btnCustom btn-sm btnEdicaoListas">Excluir</button>&nbsp;
              <button onClick={() => { this.abrirModalEditarRelatedIssues(id, titulo, arrAtribuidA, status, prioridade, txtDescricao, categoria, formdata, txtComentarios) }} className="btn btn-info btnCustom btn-sm btnEdicaoListas">Editar</button>
            </div>
          )

        }
      }

    ]

    const tablecolumnsRelatedMilestones = [
      {
        dataField: "Title",
        text: "Project Milestone",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "DueDate",
        text: "Vencimento",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var data = new Date(row.DueDate);
          var dtdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
          return dtdata;
        }
      },
      {
        dataField: "Complete",
        text: "Concluído",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var concluido = row.Complete;
          var resultado = "Não";
          if (concluido == true) resultado = "Sim";
          return resultado;
        }
      },
      {
        dataField: "ProjComments",
        text: "Comentários",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var comentarios = row.ProjComments;
          var vlrComentario = "";

          if (comentarios != null) vlrComentario = row.ProjComments;

          return <div dangerouslySetInnerHTML={{ __html: `${vlrComentario}` }} />;
        }
      },
      {
        dataField: "",
        text: "",
        headerStyle: { "backgroundColor": "#bee5eb", "width": "130px" },
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var id = row.ID;
          var titulo = row.Title;
          var concluido = row.Complete;
          var vlrConcluido = "0";
          var comentarios = row.ProjComments;

          _descricaoComentariosMilestone = comentarios;

          console.log("concluido", concluido);
          if (concluido == true) vlrConcluido = "1";

          var data = new Date(row.DueDate);
          var formdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();

          console.log("formdata", formdata);

          var txtComentarios = "";

          if (comentarios != null) {

            txtComentarios = comentarios.replace(/<[\/]{0,1}(div)[^><]*>/g, "");

            if (txtComentarios.includes("<font")) {

              txtComentarios = txtComentarios.replace("font", "span");
              txtComentarios = txtComentarios.replace("font", "span");

            }

            if (txtComentarios.includes("color")) {

              txtComentarios = txtComentarios.replace('color="', 'style="color:');

            }

            txtComentarios = txtComentarios.trim();

            console.log("txtComentarios", txtComentarios);

          }

          return (

            <div>
              <button onClick={async () => { this.excluirMilestone(id, titulo) }} className="btn btn-info btnCustom btn-sm btnEdicaoListas">Excluir</button>&nbsp;
              <button onClick={() => { this.abrirModalEditarRelatedMilestones(id, titulo, vlrConcluido, formdata, txtComentarios) }} className="btn btn-info btnCustom btn-sm btnEdicaoListas">Editar</button>
            </div>
          )

        }
      }

    ]

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
                    <div className="form-group col-md text-info ">
                      <b>Solicitação <span id='txtID'></span></b><br></br>
                      Status: <span id='txtStatus'></span>
                    </div>
                    <div className="form-group col-md text-secondary right ">

                    </div>
                  </div>
                </div>

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
            <div className="card-header btn" id="headingAnexo" data-toggle="collapse" data-target="#collapseAnexo" aria-expanded="true" aria-controls="collapseAnexo">
              <h5 className="mb-0 text-info">
                Anexos
              </h5>
            </div>
            <div id="collapseAnexo" className="collapse show" aria-labelledby="headingOne">
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

                        //var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Anexos')/items('${_projectID}')/AttachmentFiles`;
                        var url = this.props.siteurl;

                        var caminho = `${url}/Lists/Projects/Attachments/${_projectID}/${item.FileName}`;

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
            <div className="card-header btn" id="headingRelatedIssues" data-toggle="collapse" data-target="#collapseRelatedIssues" aria-expanded="true" aria-controls="collapseRelatedIssues">
              <h5 className="mb-0 text-info">
                Related Issues
              </h5>
            </div>
            <div id="collapseRelatedIssues" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">
                <div id='tabelaPreStageSoftware'>
                  <BootstrapTable bootstrap4 striped responsive condensed hover={false} className="gridTodosItens" id="gridTodosItensRelatedIssues" keyField='id' data={this.state.itemsListRelatedIssues} columns={tablecolumnsRelatedIssues} headerClasses="header-class" />
                </div>
                <button id='btnAbrirModalCadastrarRelatedIssues' className="btn btn-secondary btnCustom btn-sm">Adicionar</button>&nbsp;
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingRelatedMilestones" data-toggle="collapse" data-target="#collapseRelatedMilestones" aria-expanded="true" aria-controls="collapseRelatedMilestones">
              <h5 className="mb-0 text-info">
                Related Milestones
              </h5>
            </div>
            <div id="collapseRelatedMilestones" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">
                <div id='tabelaPreStageSoftware'>
                  <BootstrapTable bootstrap4 striped responsive condensed hover={false} className="gridTodosItens" id="gridTodosItensPreStageSoftware" keyField='id' data={this.state.itemsListRelatedMilestones} columns={tablecolumnsRelatedMilestones} headerClasses="header-class" />
                </div>
                <button id='btnAbrirModalCadastrarMilestones' className="btn btn-secondary btnCustom btn-sm">Adicionar</button>&nbsp;
              </div>
            </div>
          </div>

        </div>
      </div>

        <br></br><div className="text-right">
          <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
          <button style={{ "margin": "2px" }} id="btnConfirmarSalvar" className="btn btn-success">Salvar</button>
          <button style={{ "margin": "2px" }} id="btnConfirmarFluxo" className="btn btn-success">Iniciar fluxo</button>
        </div>

        <div className="modal fade" id="modalCadastrarRelatedIssues" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Related Issues - Cadastrar</h5>
              </div>
              <div className="modal-body">

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="txtTitulo-RelatedIssues">Título</label><span className="required"> *</span><br></br>
                    <input type="text" className="form-control" id="txtTitulo-RelatedIssues" />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="txtModeloCadastrar">Atribuido a</label><span className="required"> *</span><br></br>
                    <PeoplePicker
                      context={this.props.context as any}
                      //titleText="Aprovador Engenharia"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users
                      showtooltip={true}
                      required={true}
                      disabled={false}
                      onChange={this._getPeoplePickerAtribuidoARelatedIssues.bind(this)}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      // defaultSelectedUsers={this.state.PeoplePickerAtribudioARelatedIssues}
                      ensureUser={true} />
                  </div>
                </div>


                <div className="form-row">
                  <div className="form-group col-md">
                    <label>Status</label><span className="required"> *</span><br></br>
                    <select id="ddlStatus-RelatedIssues" className="form-control" >
                      <option value="0" selected>Selecione...</option>
                      {this.state.itemsIssueStatus.map(function (item, key) {
                        return (
                          <option value={item}>{item}</option>
                        );
                      })}
                    </select>
                  </div>
                  <div className="form-group col-md">
                    <label>Prioridade</label><br></br>
                    <select id="ddlPrioridade-RelatedIssues" className="form-control" >
                      <option value="0" selected>Selecione...</option>
                      {this.state.itemsPriority.map(function (item, key) {
                        return (
                          <option value={item}>{item}</option>
                        );
                      })}
                    </select>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="richTextDescricaoRelatedIssues">Descrição</label><br></br>
                    <div id='richTextDescricaoRelatedIssues'>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeDescricaoRelatedIssues(text)} />
                    </div>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md-9">
                    <label>Categoria</label><br></br>
                    <select id="ddlCategoria-RelatedIssues" className="form-control" >
                      <option value="0" selected>Selecione...</option>
                      {this.state.itemsIssueCategoria.map(function (item, key) {
                        return (
                          <option value={item}>{item}</option>
                        );
                      })}
                    </select>
                  </div>
                  <div className="form-group col-md-3">
                    <label htmlFor="dtData-DataVencimento-RelatedIssues">Vencimento</label><br></br>
                    <InputMask mask="99/99/9999" className="form-control" maskChar="_" id="dtData-DataVencimento-RelatedIssues" />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="richTextComentariosRelatedIssues">Comentários</label><br></br>
                    <div id='richTextComentariosRelatedIssues'>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeComentariosRelatedIssues(text)} />
                    </div>
                  </div>
                </div>

              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnCadastrarRelatedIssues" className="btn btn-success">Salvar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalEditarRelatedIssues" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Related Issues - Editar</h5>
              </div>
              <div className="modal-body">

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="txtTitulo-RelatedIssues-Editar">Título</label><span className="required"> *</span><br></br>
                    <input type="text" className="form-control" id="txtTitulo-RelatedIssues-Editar" />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="txtModeloCadastrar">Atribuido a</label><span className="required"> *</span><br></br>
                    <PeoplePicker
                      context={this.props.context as any}
                      //titleText="Aprovador Engenharia"
                      personSelectionLimit={1}
                      groupName={""} // Leave this blank in case you want to filter from all users
                      showtooltip={true}
                      required={true}
                      disabled={false}
                      onChange={this._getPeoplePickerAtribuidoARelatedIssues.bind(this)}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      defaultSelectedUsers={this.state.PeoplePickerAtribudioARelatedIssuesEditar}
                      ensureUser={true} />
                  </div>
                </div>


                <div className="form-row">
                  <div className="form-group col-md">
                    <label>Status</label><span className="required"> *</span><br></br>
                    <select id="ddlStatus-RelatedIssues-Editar" className="form-control" >
                      <option value="0" selected>Selecione...</option>
                      {this.state.itemsIssueStatus.map(function (item, key) {
                        return (
                          <option value={item}>{item}</option>
                        );
                      })}
                    </select>
                  </div>
                  <div className="form-group col-md">
                    <label>Prioridade</label><br></br>
                    <select id="ddlPrioridade-RelatedIssues-Editar" className="form-control" >
                      <option value="0" selected>Selecione...</option>
                      {this.state.itemsPriority.map(function (item, key) {
                        return (
                          <option value={item}>{item}</option>
                        );
                      })}
                    </select>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="richTextDescricaoRelatedIssues">Descrição</label><br></br>
                    <div id='richTextDescricaoRelatedIssues-Editar'>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeDescricaoRelatedIssues(text)} />
                    </div>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md-9">
                    <label>Categoria</label><br></br>
                    <select id="ddlCategoria-RelatedIssues-Editar" className="form-control" >
                      <option value="0" selected>Selecione...</option>
                      {this.state.itemsIssueCategoria.map(function (item, key) {
                        return (
                          <option value={item}>{item}</option>
                        );
                      })}
                    </select>
                  </div>
                  <div className="form-group col-md-3">
                    <label htmlFor="dtData-DataVencimento-RelatedIssues-Editar">Vencimento</label><br></br>
                    <InputMask mask="99/99/9999" className="form-control" value={this.state.valorItemsDataVencimentoRelatedIssues} maskChar="_" id="dtData-DataVencimento-RelatedIssues-Editar" onChange={(e) => this.onTextChangeDataVencimentoRelatedIssues(e.target.value)} />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label>Comentários</label><br></br>
                    <div id='richTextComentariosRelatedIssues-Editar'>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeComentariosRelatedIssues(text)} />
                    </div>
                  </div>
                </div>

              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnEditarRelatedIssues" className="btn btn-success">Salvar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalCadastrarMilestone" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Related Milestones - Cadastrar</h5>
              </div>
              <div className="modal-body">

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="txtProjectMilestone">Project Milestone</label><span className="required"> *</span><br></br>
                    <InputMask mask="(99999) aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa" className="form-control" maskChar="" id="txtProjectMilestone" />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md-9">
                    <label>Concluído</label><br></br>
                    <select id="ddlConcluido-Milestone" className="form-control" >
                      <option value="0" selected>Não</option>
                      <option value="1" >Sim</option>
                    </select>
                  </div>
                  <div className="form-group col-md-3">
                    <label htmlFor="dtData-DataVencimento-Milestone">Vencimento</label><br></br>
                    <InputMask mask="99/99/9999" className="form-control" maskChar="_" id="dtData-DataVencimento-Milestone" />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="richTextComentariosRelatedIssues">Comentários</label><br></br>
                    <div id='richTextComentariosRelatedIssues'>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeComentariosMilestone(text)} />
                    </div>
                  </div>
                </div>


              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnCadastrarMilestone" className="btn btn-success">Salvar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalEditarMilestone" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Related Milestones - Editar</h5>
              </div>
              <div className="modal-body">

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="txtProjectMilestone-Editar">Project Milestone</label><span className="required"> *</span><br></br>
                    <InputMask mask="(99999) aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa" className="form-control" maskChar="" id="txtProjectMilestone-Editar" value={this.state.valorProjectMilestoneRelatedMilestones} onChange={(e) => this.onTextChangeProjectMilestoneRelatedMilestones(e.target.value)} />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md-9">
                    <label>Concluído</label><br></br>
                    <select id="ddlConcluido-Milestone-Editar" className="form-control" >
                      <option value="0" selected>Não</option>
                      <option value="1" >Sim</option>
                    </select>
                  </div>
                  <div className="form-group col-md-3">
                    <label htmlFor="dtData-DataVencimento-RelatedMilestone-Editar">Vencimento</label><br></br>
                    <InputMask mask="99/99/9999" className="form-control" maskChar="_" id="dtData-DataVencimento-Milestone-Editar" value={this.state.valorItemsDataVencimentoRelatedMilestones} onChange={(e) => this.onTextChangeDataVencimentoRelatedMilestones(e.target.value)} />
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md">
                    <label htmlFor="richTextComentariosRelatedMilestoneEditar">Comentários</label><br></br>
                    <div id='richTextComentariosMilestone-Editar'>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeComentariosMilestone(text)} />
                    </div>
                  </div>
                </div>


              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnEditarMilestone" className="btn btn-success">Salvar</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoCadastrarRelatedIssues" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Related Issue criado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoCadastrarRelatedIssues" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoEditarRelatedIssues" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Related Issue alterado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoEditarRelatedIssues" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoCadastrarMilestone" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Related Milestone criado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoCadastrarMilestone" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoEditarMilestone" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Related Milestone alterado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoEditarMilestone" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoExcluirRelatedIssue" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Related Issue excluido com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoExcluirRelatedIssues" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>


        <div className="modal fade" id="modalSucessoExcluirMilestone" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Related Milestone excluido com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoExcluirMilestone" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalConfirmarSalvar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
              </div>
              <div className="modal-body">
                Deseja realmente salvar a Solicitação?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnSalvar" type="button" className="btn btn-primary">Sim</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalConfirmarAprovar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
              </div>
              <div className="modal-body">
                Deseja realmente iniciar o Fluxo?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnIniciarFluxo" type="button" className="btn btn-primary">Sim</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoSalvar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Solicitação alterada com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoSalvar" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoAprovar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Fluxo iniciado com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoAprovar" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalCarregando" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div>
            <div className="modal-dialog" role="document">
              <div className="modal-content">
                <div id='conteudoLoading' className='carregando'></div>
              </div>
            </div>
          </div>
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


    var reactIssueStatus = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('Project Issues')/fields?$filter=EntityPropertyName eq 'Status'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactIssueStatus.setState({
          itemsIssueStatus: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    var reactPriority = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('Project Issues')/fields?$filter=EntityPropertyName eq 'Priority'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactPriority.setState({
          itemsPriority: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    var reactIssueCategoria = this;

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('Project Issues')/fields?$filter=EntityPropertyName eq 'Category'`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactIssueCategoria.setState({
          itemsIssueCategoria: resultData.d.results[0].Choices.results
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    var reactHandlerRelatedIssues = this;

    if (_siteNovo) {

      var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Issues')/items?$top=50&$orderby= Created asc&$select=ID,Title,Priority,Status,AssignedTo/ID,AssignedTo/Title,DueDate,Comment,Category,V3Comments,Assigned_x0020_To_x0020_2&$expand=AssignedTo&$filter=Project/ID eq ` + _projectID;
    }
    else {

      var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Issues')/items?$top=50&$orderby= Created asc&$select=ID,Title,Priority,Status,AssignedTo/ID,AssignedTo/Title,DueDate,Comment,Category,V3Comments,Assigned_x0020_To_x0020_2&$expand=AssignedTo&$filter=Project_x0020_TXT eq '${_projectTitle}'`;

    }


    jQuery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        if (resultData.d.results.length > 0) {
          jQuery("#tabelaPreStageSoftware").show();
          reactHandlerRelatedIssues.setState({
            itemsListRelatedIssues: resultData.d.results
          });
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    var reactHandlerRelatedMilestones = this;

    if (_siteNovo) {

      var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Milestones')/items?$top=50&$orderby= Created asc&$select=ID,Title,Complete,DueDate,ProjComments&$filter=Project/ID eq ` + _projectID;
    }
    else {

      var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Milestones')/items?$top=50&$orderby= Created asc&$select=ID,Title,Complete,DueDate,ProjComments&$filter=Project_x0020_TXT eq '${_projectTitle}'`;

    }

    jQuery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        if (resultData.d.results.length > 0) {
          jQuery("#tabelaPreStageSoftware").show();
          reactHandlerRelatedMilestones.setState({
            itemsListRelatedMilestones: resultData.d.results
          });
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });





  }

  protected getProject() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Projects List')/items?$select=ID,Title,ProjCategory,Project_x0020_type,AssignedTo/ID,AssignedTo/Title,Participants/ID,Participants/Title,Product_x0020_description_x0020_,Critical_x0020_requirements,Client/ID,OMP_x0020_documents,ProjStatus,SiteNovoSO,Owner_x0020_2,Participants_x0020_2&$expand=AssignedTo,Participants,Client&$filter=ID eq ` + _projectID,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        //  console.log("resultData doc", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var id = resultData.d.results[i].ID;
            var status = resultData.d.results[i].ProjStatus;
            var title = resultData.d.results[i].Title;

            _projectTitle = title;

            _statusAtual = status;

            jQuery("#txtID").html(id);
            jQuery("#txtStatus").html(status);

            var nome = resultData.d.results[i].Title;
            var category = resultData.d.results[i].ProjCategory;
            var tipo = resultData.d.results[i].Project_x0020_type;
            var omp = resultData.d.results[i].OMP_x0020_documents;

            var siteNovo = resultData.d.results[i].SiteNovoSO;
            _siteNovo = siteNovo;

            console.log("siteNovo", siteNovo);

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


            if ((status == "Não Iniciada") || (status == "Em Andamento") || (status == "Adiada")) {

              if (_grupos.indexOf("SST - Elaboradores") !== -1) {

                jQuery("#btnConfirmarSalvar").show();
                jQuery("#btnAbrirModalCadastrarRelatedIssues").show();
                jQuery("#btnAbrirModalCadastrarMilestones").show();
                jQuery(".btnEdicaoListas").show();

                if ((status == "Não Iniciada") || (status == "Adiada")) {

                  setTimeout(() => {

                    jQuery("#btnConfirmarFluxo").show();
  
                  }, 1000);
                  
                }

              } else {

                setTimeout(() => {

                  jQuery(".btnEdicaoListas").hide();

                }, 1000);

              }

              

            } else {


              setTimeout(() => {

                jQuery(".btnEdicaoListas").hide();

              }, 1000);

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

  protected abrirModalCadastrarRelatedIssues() {

    console.log("PeoplePickerAtribudioARelatedIssues", this.state.PeoplePickerAtribudioARelatedIssues);

    jQuery("#txtTitulo-RelatedIssues").val("");
    jQuery("#ddlStatus-RelatedIssues").val("0");
    jQuery("#ddlPrioridade-RelatedIssues").val("0");
    jQuery('#richTextDescricaoRelatedIssues').find('.ql-editor').html("<p><br></p>");
    jQuery("#ddlCategoria-RelatedIssues").val("0");
    jQuery("#dtData-DataVencimento-RelatedIssues").val("");
    jQuery('#richTextComentariosRelatedIssues').find('.ql-editor').html("<p><br></p>");


    jQuery("#modalCadastrarRelatedIssues").modal({ backdrop: 'static', keyboard: false });



  }

  protected abrirModalCadastrarMilestones() {


    jQuery("#modalCadastrarMilestone").modal({ backdrop: 'static', keyboard: false });



  }

  protected async cadastrarProjectIssues() {

    jQuery("#btnCadastrarRelatedIssues").prop("disabled", true);

    var titulo = jQuery("#txtTitulo-RelatedIssues").val();
    var status = jQuery("#ddlStatus-RelatedIssues option:selected").val();
    var prioridade = jQuery("#ddlPrioridade-RelatedIssues option:selected").val();
    var descricao = _descricaoRelatedIssues;
    var categoria = jQuery("#ddlCategoria-RelatedIssues option:selected").val();
    var comentario = _descricaoComentariosRelatedIssues;

    var data = "" + $("#dtData-DataVencimento-RelatedIssues").val() + "";
    var dataDia = data.substring(0, 2);
    var dataMes = data.substring(3, 5);
    var dataAno = data.substring(6, 10);
    var formData = dataAno + "-" + dataMes + "-" + dataDia;

    var arrAtribuidoARelatedIssues = [];
    for (let i = 0; i < _atribuidoARelatedIssues.length; i++) {
      arrAtribuidoARelatedIssues.push(_atribuidoARelatedIssues[i]["id"]);
    }

    //validacao
    if (titulo == "") {
      alert("Forneça o Título!");
      jQuery("#btnCadastrarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (_atribuidoARelatedIssues.length == 0) {
      alert("Forneça pra quem será atribuído!");
      jQuery("#btnCadastrarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (status == "0") {
      alert("Forneça o Status!");
      jQuery("#btnCadastrarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (titulo == "") {
      alert("Forneça o Título!");
      jQuery("#btnCadastrarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (data == "") {
      data = null;
    } else {
      var reg = /(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d/;
      if (data.match(reg)) {
      }
      else {
        alert("Forneça uma data válida!");
        jQuery("#btnCadastrarPontoCorte").prop("disabled", false);
        return false;
      }
    }

    //cadastrar

    await _web.lists
      .getByTitle("Project Issues")
      .items.add({
        ProjectId: _projectID,
        Title: titulo,
        Status: status,
        Priority: prioridade,
        Comment: descricao,
        Category: categoria,
        DueDate: formData,
        V3Comments: comentario,
        AssignedToId: { 'results': arrAtribuidoARelatedIssues },
      })
      .then(response => {

        jQuery("#btnCadastrarRelatedIssues").prop("disabled", false);
        jQuery("#modalCadastrarRelatedIssues").modal('hide');
        jQuery("#modalSucessoCadastrarRelatedIssues").modal({ backdrop: 'static', keyboard: false });


      })
      .catch((error: any) => {
        console.log(error);
      })



  }

  protected async abrirModalEditarRelatedIssues(id, titulo, arrAtribuidA, status, prioridade, descricao, categoria, vencimento, txtComentarios) {

    jQuery("#btnEditarRelatedIssues").prop("disabled", false);
    // 
    //console.log("arrValAtribuidAID x", arrValAtribuidAID);

    // _arrValAtribuidAID = arrValAtribuidAID;
    // _arrValAtribuidAID = arrValAtribuidAID;

    _idRelatedIssues = id;

    var arrValAtribuidA = [];
    var arrValAtribuidAID = [];

    _arrValAtribuidAID = [];

    console.log("arrAtribuidA A", arrAtribuidA);

    if (arrAtribuidA.hasOwnProperty('results')) {
      for (var x = 0; x < arrAtribuidA.results.length; x++) {
        arrValAtribuidA.push(arrAtribuidA.results[x].Title);
        arrValAtribuidAID.push(arrAtribuidA.results[x].ID);
      }
    }

    console.log("arrValAtribuidA A", arrValAtribuidA);
    console.log("arrValAtribuidAID A", arrValAtribuidAID);

    _arrValAtribuidAID = arrValAtribuidAID;

    jQuery('#txtTitulo-RelatedIssues-Editar').val(titulo);
    jQuery(`#ddlStatus-RelatedIssues-Editar option[value="${status}"]`).prop('selected', true);
    jQuery(`#ddlPrioridade-RelatedIssues-Editar option[value="${prioridade}"]`).prop('selected', true);
    jQuery('#richTextDescricaoRelatedIssues-Editar').find('.ql-editor').html(`${descricao}`);
    jQuery(`#ddlCategoria-RelatedIssues-Editar option[value="${categoria}"]`).prop('selected', true);
    jQuery('#richTextComentariosRelatedIssues-Editar').find('.ql-editor').html(`${txtComentarios}`);

    var reactRelatedIssues = this;
    reactRelatedIssues.setState({
      valorItemsDataVencimentoRelatedIssues: vencimento
    });

    // console.log("atribuidA", atribuidA);

    setTimeout(() => {

      this.setState({
        PeoplePickerAtribudioARelatedIssuesEditar: arrValAtribuidA,
      });

    }, 0);

    jQuery("#modalEditarRelatedIssues").modal({ backdrop: 'static', keyboard: false });

  }

  protected async editarProjectIssues() {

    jQuery("#btnEditarRelatedIssues").prop("disabled", true);

    var titulo = jQuery("#txtTitulo-RelatedIssues-Editar").val();
    var status = jQuery("#ddlStatus-RelatedIssues-Editar option:selected").val();
    var prioridade = jQuery("#ddlPrioridade-RelatedIssues-Editar option:selected").val();
    var descricao = _descricaoRelatedIssues;
    var categoria = jQuery("#ddlCategoria-RelatedIssues-Editar option:selected").val();
    var comentario = _descricaoComentariosRelatedIssues;

    var data = "" + $("#dtData-DataVencimento-RelatedIssues-Editar").val() + "";
    var dataDia = data.substring(0, 2);
    var dataMes = data.substring(3, 5);
    var dataAno = data.substring(6, 10);
    var formData = dataAno + "-" + dataMes + "-" + dataDia;

    var arrAtribuidoARelatedIssues = [];

    console.log("_atribuidoARelatedIssuesTeveAlteracao", _atribuidoARelatedIssuesTeveAlteracao);
    console.log("_arrValAtribuidAID", _arrValAtribuidAID);

    if (_atribuidoARelatedIssuesTeveAlteracao) {

      for (let i = 0; i < _atribuidoARelatedIssues.length; i++) {
        if (_atribuidoARelatedIssues[i]["id"] != null)
          arrAtribuidoARelatedIssues.push(_atribuidoARelatedIssues[i]["id"]);
      }

    } else {

      arrAtribuidoARelatedIssues = _arrValAtribuidAID;

    }

    console.log("arrAtribuidoARelatedIssues", arrAtribuidoARelatedIssues);


    //validacao
    if (titulo == "") {
      alert("Forneça o Título!");
      jQuery("#btnEditarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (_atribuidoARelatedIssuesTeveAlteracao) {

      if (_atribuidoARelatedIssues.length == 0) {
        alert("Forneça pra quem será atribuído!");
        jQuery("#btnEditarRelatedIssues").prop("disabled", false);
        return false;
      }

    }

    if (status == "0") {
      alert("Forneça o Status!");
      jQuery("#btnEditarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (titulo == "") {
      alert("Forneça o Título!");
      jQuery("#btnEditarRelatedIssues").prop("disabled", false);
      return false;
    }

    if (data == "") {
      data = null;
    } else {
      var reg = /(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d/;
      if (data.match(reg)) {
      }
      else {
        alert("Forneça uma data válida!");
        jQuery("#btnCadastrarPontoCorte").prop("disabled", false);
        return false;
      }
    }

    //cadastrar

    await _web.lists
      .getByTitle("Project Issues")
      .items.getById(_idRelatedIssues).update({
        ProjectId: _projectID,
        Title: titulo,
        Status: status,
        Priority: prioridade,
        Comment: descricao,
        Category: categoria,
        DueDate: formData,
        V3Comments: comentario,
        AssignedToId: { 'results': arrAtribuidoARelatedIssues },
      })
      .then(response => {

        console.log("editou!!!");

        jQuery("#btnEditarRelatedIssues").prop("disabled", false);
        jQuery("#modalEditarRelatedIssues").modal('hide');
        jQuery("#modalSucessoEditarRelatedIssues").modal({ backdrop: 'static', keyboard: false });


      })
      .catch((error: any) => {
        console.log(error);
      })



  }

  protected async cadastrarMilestone() {

    jQuery("#btnCadastrarMilestone").prop("disabled", true);

    var projectMilestone = jQuery("#txtProjectMilestone").val();
    var concluido = jQuery("#ddlConcluido-Milestone option:selected").val();
    var complete = false;

    if (concluido == "1") complete = true;


    var data = "" + $("#dtData-DataVencimento-Milestone").val() + "";
    var dataDia = data.substring(0, 2);
    var dataMes = data.substring(3, 5);
    var dataAno = data.substring(6, 10);
    var formData = dataAno + "-" + dataMes + "-" + dataDia;

    var comentario = _descricaoComentariosMilestone;

    //validacao
    if (projectMilestone == "") {
      alert("Forneça o nome!");
      jQuery("#btnCadastrarMilestone").prop("disabled", false);
      return false;
    }

    if (data == "") {
      data = null;
    } else {
      var reg = /(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d/;
      if (data.match(reg)) {
      }
      else {
        alert("Forneça uma data válida!");
        jQuery("#btnCadastrarPontoCorte").prop("disabled", false);
        return false;
      }
    }

    //cadastrar

    await _web.lists
      .getByTitle("Project Milestones")
      .items.add({
        ProjectId: _projectID,
        Title: projectMilestone,
        Complete: complete,
        DueDate: formData,
        ProjComments: comentario

      })
      .then(response => {

        jQuery("#btnCadastrarMilestone").prop("disabled", false);
        jQuery("#modalCadastrarMilestone").modal('hide');
        jQuery("#modalSucessoCadastrarMilestone").modal({ backdrop: 'static', keyboard: false });


      })
      .catch((error: any) => {
        console.log(error);
      })



  }

  protected async editarMilestone() {

    jQuery("#btnEditarMilestone").prop("disabled", true);

    var projectMilestone = jQuery("#txtProjectMilestone-Editar").val();
    var concluido = jQuery("#ddlConcluido-Milestone-Editar option:selected").val();
    var complete = false;

    if (concluido == "1") complete = true;

    var data = "" + $("#dtData-DataVencimento-Milestone-Editar").val() + "";
    var dataDia = data.substring(0, 2);
    var dataMes = data.substring(3, 5);
    var dataAno = data.substring(6, 10);
    var formData = dataAno + "-" + dataMes + "-" + dataDia;

    var comentario = _descricaoComentariosMilestone;

    //validacao
    if (projectMilestone == "") {
      alert("Forneça o nome!");
      jQuery("#btnEditarMilestone").prop("disabled", false);
      return false;
    }

    if (data == "") {
      data = null;
    } else {
      var reg = /(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d/;
      if (data.match(reg)) {
      }
      else {
        alert("Forneça uma data válida!");
        jQuery("#btnCadastrarPontoCorte").prop("disabled", false);
        return false;
      }
    }

    //cadastrar

    await _web.lists
      .getByTitle("Project Milestones")
      .items.getById(_idRelatedMilestones).update({
        ProjectId: _projectID,
        Title: projectMilestone,
        Complete: complete,
        DueDate: formData,
        ProjComments: comentario

      })
      .then(response => {

        jQuery("#btnEditarMilestone").prop("disabled", false);
        jQuery("#modalEditarMilestone").modal('hide');
        jQuery("#modalSucessoEditarMilestone").modal({ backdrop: 'static', keyboard: false });


      })
      .catch((error: any) => {
        console.log(error);
      })



  }

  protected fecharSucessoRelatedIssues() {

    var reactHandlerRelatedIssues = this;

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Project Issues')/items?$top=50&$orderby= Created asc&$select=ID,Title,Priority,Status,AssignedTo/ID,AssignedTo/Title,DueDate,Comment,Category,V3Comments&$expand=AssignedTo&$filter=Project/ID eq ` + _projectID,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        if (resultData.d.results.length > 0) {
          //jQuery("#tabelaPreStageSoftware").show();
          reactHandlerRelatedIssues.setState({
            itemsListRelatedIssues: resultData.d.results
          });
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    $("#modalSucessoCadastrarRelatedIssues").modal('hide');
    $("#modalSucessoExcluirRelatedIssue").modal('hide');
    $("#modalSucessoEditarRelatedIssues").modal('hide');


  }

  protected fecharSucessoRelatedMilestone() {

    var reactHandlerRelatedMilestones = this;

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Project Milestones')/items?$top=50&$orderby= Created asc&$select=ID,Title,Complete,DueDate,ProjComments&$filter=Project/ID eq ` + _projectID,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData", resultData);

        if (resultData.d.results.length > 0) {
          jQuery("#tabelaPreStageSoftware").show();
          reactHandlerRelatedMilestones.setState({
            itemsListRelatedMilestones: resultData.d.results
          });
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });

    $("#modalSucessoCadastrarMilestone").modal('hide');
    $("#modalSucessoExcluirMilestone").modal('hide');
    $("#modalSucessoEditarMilestone").modal('hide');

  }

  protected async excluirRelatedIssues(id, titulo) {

    if (confirm("Deseja realmente excluir o Related Issue: " + titulo + "?") == true) {

      const list = _web.lists.getByTitle("Project Issues");
      await list.items.getById(id).recycle()
        .then(async response => {

          console.log("excluido");
          jQuery("#modalSucessoExcluirRelatedIssue").modal({ backdrop: 'static', keyboard: false });


        })
        .catch((error: any) => {
          console.log(error);

        })


    } else {

      return false.valueOf;
    }

  }

  protected async abrirModalEditarRelatedMilestones(id, titulo, concluido, vencimento, txtComentarios) {

    console.log("titulo", titulo);

    _idRelatedMilestones = id;

    //jQuery('#txtProjectMilestone-Editar').val(titulo);
    jQuery(`#ddlConcluido-Milestone-Editar option[value="${concluido}"]`).prop('selected', true);

    var reactRelatedIssues = this;
    reactRelatedIssues.setState({
      valorItemsDataVencimentoRelatedMilestones: vencimento,
      valorProjectMilestoneRelatedMilestones: titulo
    });

    jQuery('#richTextComentariosMilestone-Editar').find('.ql-editor').html(`${txtComentarios}`);


    jQuery("#modalEditarMilestone").modal({ backdrop: 'static', keyboard: false });

  }

  protected async excluirMilestone(id, titulo) {

    if (confirm("Deseja realmente excluir o Related Milestone: " + titulo + "?") == true) {

      const list = _web.lists.getByTitle("Project Milestones");
      await list.items.getById(id).recycle()
        .then(async response => {

          console.log("excluido");
          jQuery("#modalSucessoExcluirMilestone").modal({ backdrop: 'static', keyboard: false });


        })
        .catch((error: any) => {
          console.log(error);

        })


    } else {

      return false.valueOf;
    }

  }

  protected validar(opcao) {

    var name = jQuery("#txtName").val();
    var category = jQuery("#ddlCategory option:selected").val();
    var type = jQuery("#ddlType option:selected").val();
    var ompDocuments = jQuery("#txtOMPDocuments").val();

    var arrClient = Array.prototype.slice.call(document.querySelectorAll('#ddlCliente2 option'), 0).map(function (v, i, a) {
      return v.value;
    });

    if (name == "") {
      alert("Forneça o nome da Solicitação!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    if (category == "0") {
      alert("Escolha uma categoria!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    if (type == "0") {
      alert("Escolha um tipo!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    if (_arrOwner.length == 0) {
      alert("Forneça o Owner!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    if (_arrParticipants.length == 0) {
      alert("Forneça os participantes!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    if (arrClient.length == 0) {
      alert("Escolha pelo meno um cliente!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    if (ompDocuments == "") {
      alert("Forneça a OMP!");
      document.getElementById('ProjectInformation').scrollIntoView();
      return false;
    }

    var files = (document.querySelector("#input") as HTMLInputElement).files;

    if (files.length > 0) {

      console.log("files.length", files.length);

      for (var i = 0; i <= files.length - 1; i++) {

        var fsize = files.item(i).size;
        _size = _size + fsize;

        console.log("fsize", fsize);

      }

      if (_size > 15000000) {
        alert("A soma dos arquivos não pode ser maior que 15mega!");
        _size = 0;
        return false;
      }

    }

    if (opcao == "Salvar") jQuery("#modalConfirmarSalvar").modal({ backdrop: 'static', keyboard: false });
    if (opcao == "Aprovar") jQuery("#modalConfirmarAprovar").modal({ backdrop: 'static', keyboard: false });

  }

  protected async editar(opcao) {

    jQuery("#btnSalvar").prop("disabled", true);

    jQuery("#modalConfirmarSalvar").modal('hide');
    jQuery("#modalConfirmarAprovar").modal('hide');
    jQuery("#modalCarregando").modal({ backdrop: 'static', keyboard: false });

    var name = jQuery("#txtName").val();
    var category = jQuery("#ddlCategory option:selected").val();
    var type = jQuery("#ddlType option:selected").val();
    var ompDocuments = jQuery("#txtOMPDocuments").val();

    if (_ownerTevelateracao) {

      var arrPeoplepickerOwner = [];
      for (let i = 0; i < _arrOwner.length; i++) {
        console.log('_arrOwner[i]["id"]', _arrOwner[i]["id"]);
        arrPeoplepickerOwner.push(_arrOwner[i]["id"]);
      }

    } else {
      arrPeoplepickerOwner = _arrOwnerID;
    }

    if (_participantsTevelateracao) {

      var arrPeoplepickerParticipants = [];
      for (let i = 0; i < _arrParticipants.length; i++) {
        arrPeoplepickerParticipants.push(_arrParticipants[i]["id"]);
      }

    }

    else {

      arrPeoplepickerParticipants = _arrParticipantsID;

    }

    var arrClient = Array.prototype.slice.call(document.querySelectorAll('#ddlCliente2 option'), 0).map(function (v, i, a) {
      return v.value;
    });

    console.log("_arrOwner 1", _arrOwner);
    console.log("arrPeoplepickerOwner 1", arrPeoplepickerOwner);

    var status;

    if (opcao == "Aprovar") status = "Em Andamento";
    else status = _statusAtual;

    await _web.lists
      .getByTitle("Projects List")
      .items.getById(_projectID).update({
        Title: name,
        ProjCategory: category,
        Project_x0020_type: type,
        AssignedToId: { 'results': arrPeoplepickerOwner },
        ParticipantsId: { 'results': arrPeoplepickerParticipants },
        Product_x0020_description_x0020_: _productDescription,
        Critical_x0020_requirements: _criticalRequirements,
        ClientId: { "results": arrClient },
        OMP_x0020_documents: ompDocuments,
        ProjStatus: status
      })
      .then(response => {
        console.log("editou!!");
        this.upload(opcao);
      })
      .catch((error: any) => {
        console.log(error);
      })

  }

  protected async upload(opcao) {

    console.log("Entrou no upload");

    var files = (document.querySelector("#input") as HTMLInputElement).files;
    var file = files[0];

    //console.log("files.length", files.length);

    if (files.length != 0) {

      if (_pastaCriada != "Sim") {

        _web.lists.getByTitle("Anexos").rootFolder.folders.add(`${_projectID}`).then(async data => {

          await _web.lists
            .getByTitle("Projects List")
            .items.getById(_projectID).update({
              PastaCriada: "Sim",
            })
            .then(async response => {

              for (var i = 0; i < files.length; i++) {

                var nomeArquivo = files[i].name;
                var rplNomeArquivo = nomeArquivo.replace(/[^0123456789.,a-zA-Z]/g, '');

                //alert(rplNomeArquivo);
                //Upload a file to the SharePoint Library
                _web.getFolderByServerRelativeUrl(`${_caminho}/Anexos/${_projectID}`)
                  //.files.add(files[i].name, files[i], true)
                  .files.add(rplNomeArquivo, files[i], true)
                  .then(async data => {

                    if (i == files.length) {
                      console.log("anexou:" + rplNomeArquivo);
                      jQuery("#modalCarregando").modal('hide');
                      if (opcao == "Salvar") jQuery("#modalSucessoSalvar").modal({ backdrop: 'static', keyboard: false });
                      if (opcao == "Aprovar") jQuery("#modalSucessoAprovar").modal({ backdrop: 'static', keyboard: false });
                    }
                  });
              }


            }).catch(err => {
              console.log("err", err);
            });



        }).catch(err => {
          console.log("err", err);
        });

      }
      else {

        await _web.lists
          .getByTitle("Projects List")
          .items.getById(_projectID).update({
            PastaCriada: "Sim",
          })
          .then(async response => {

            for (var i = 0; i < files.length; i++) {

              var nomeArquivo = files[i].name;
              var rplNomeArquivo = nomeArquivo.replace(/[^0123456789.,a-zA-Z]/g, '');

              //alert(rplNomeArquivo);
              //Upload a file to the SharePoint Library
              _web.getFolderByServerRelativeUrl(`${_caminho}/Anexos/${_projectID}`)
                //.files.add(files[i].name, files[i], true)
                .files.add(rplNomeArquivo, files[i], true)
                .then(async data => {

                  if (i == files.length) {
                    console.log("anexou:" + rplNomeArquivo);
                    jQuery("#modalCarregando").modal('hide');
                    if (opcao == "Salvar") jQuery("#modalSucessoSalvar").modal({ backdrop: 'static', keyboard: false });
                    if (opcao == "Aprovar") jQuery("#modalSucessoAprovar").modal({ backdrop: 'static', keyboard: false });
                  }
                });
            }


          }).catch(err => {
            console.log("err", err);
          });



        //const folderAddResult = _web.folders.add(`${_caminho}/Anexos/${_idProposta}`);
        //console.log("foi");
      }

    } else {

      if (_pastaCriada != "Sim") {

        _web.lists.getByTitle("Anexos").rootFolder.folders.add(`${_projectID}`).then(async data => {

          console.log("Gravou!!");
          jQuery("#modalCarregando").modal('hide');
          if (opcao == "Salvar") jQuery("#modalSucessoSalvar").modal({ backdrop: 'static', keyboard: false });
          if (opcao == "Aprovar") jQuery("#modalSucessoAprovar").modal({ backdrop: 'static', keyboard: false });

        }).catch(err => {
          console.log("err", err);
        });

      } else {

        console.log("Gravou!!");
        jQuery("#modalCarregando").modal('hide');
        if (opcao == "Salvar") jQuery("#modalSucessoSalvar").modal({ backdrop: 'static', keyboard: false });
        if (opcao == "Aprovar") jQuery("#modalSucessoAprovar").modal({ backdrop: 'static', keyboard: false });

      }

    }



  }


  protected async fecharSucessoEditar(opcao) {

    jQuery("#modalSucessoSalvar").modal('hide');
    jQuery("#modalSucessoAprovar").modal('hide');

    if (opcao == "Salvar") window.location.href = `Solicitacao-Editar.aspx?ProjectID=` + _projectID;
    if (opcao == "Aprovar") window.location.href = `Solicitacao-Todos.aspx`;

  }


  private getDefaultUsers() {

    setTimeout(() => {

      console.log("_arrOwner", _arrOwner);

      this.setState({
        PeoplePickerDefaultItemsOwner: _arrOwner,
        PeoplePickerDefaultItemsParticipants: _arrParticipants,

      });

    }, 500);


  }

  private _getPeoplePickerOwner(items: any[]) {
    _ownerTevelateracao = true;
    _arrOwner = items;
    console.log('Items:', items);
  }

  private _getPeoplePickerAtribuidoARelatedIssues(items: any[]) {
    _atribuidoARelatedIssuesTeveAlteracao = true;
    _atribuidoARelatedIssues = items;
    console.log('Items:', items);
  }

  private _getPeoplePickerParticipants(items: any[]) {
    _participantsTevelateracao = true;
    _arrParticipants = items;
    console.log('Items:', items);
  }

  private onTextChangeProductDescription = (newText: string) => {
    _productDescription = newText;
    return newText;
  }

  private onTextChangeDescricaoRelatedIssues = (newText: string) => {
    _descricaoRelatedIssues = newText;
    return newText;
  }

  private onTextChangeComentariosRelatedIssues = (newText: string) => {
    _descricaoComentariosRelatedIssues = newText;
    return newText;
  }

  private onTextChangeComentariosMilestone = (newText: string) => {
    _descricaoComentariosMilestone = newText;
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

  private onTextChangeDataVencimentoRelatedIssues = (val) => {

    this.setState({
      valorItemsDataVencimentoRelatedIssues: val,
    });

  }

  private onTextChangeDataVencimentoRelatedMilestones = (val) => {

    this.setState({
      valorItemsDataVencimentoRelatedMilestones: val,
    });

  }

  private onTextChangeProjectMilestoneRelatedMilestones = (val) => {

    this.setState({
      valorProjectMilestoneRelatedMilestones: val,
    });

  }

  protected voltar() {
    history.back();
  }


}
