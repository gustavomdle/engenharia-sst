import * as React from 'react';
import styles from './SstDetalhes.module.scss';
import { ISstDetalhesProps } from './ISstDetalhesProps';
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
var _correntUser;
var _projectTitle;
var _siteNovo;

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


export default class SstDetalhes extends React.Component<ISstDetalhesProps, IReactGetItemsState> {

  public constructor(props: ISstDetalhesProps, state: IReactGetItemsState) {
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
      .getElementById("btnVoltar")
      .addEventListener("click", (e: Event) => this.voltar());

    document
      .getElementById("btnConfirmarAdiar")
      .addEventListener("click", (e: Event) => this.confirmarEditar("Adiar"));

    document
      .getElementById("btnConfirmarCancelar")
      .addEventListener("click", (e: Event) => this.confirmarEditar("Cancelar"));

    document
      .getElementById("btnConfirmarConcluir")
      .addEventListener("click", (e: Event) => this.confirmarEditar("Concluir"));

    document
      .getElementById("btnEditar")
      .addEventListener("click", (e: Event) => this.editar("Editar"));

    document
      .getElementById("btnAdiar")
      .addEventListener("click", (e: Event) => this.editar("Adiar"));

    document
      .getElementById("btnCancelar")
      .addEventListener("click", (e: Event) => this.editar("Cancelar"));

    document
      .getElementById("btnConcluir")
      .addEventListener("click", (e: Event) => this.editar("Concluir"));

    document
      .getElementById("btnSucessoAdiar")
      .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

    document
      .getElementById("btnSucessoCancelar")
      .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

    document
      .getElementById("btnSucessoConcluir")
      .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

    jQuery("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/SiteAssets/loading.gif'/>
      <br/>Aguarde....<br/><br/>
      Dependendo do tamanho do anexo e a velocidade<br>
       da Internet essa ação pode demorar um pouco. <br>
       Não fechar a janela!<br/><br/>`);

    await _web.currentUser.get().then(f => {
      console.log("user", f);
      var id = f.Id;
      _correntUser = f.Title;

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




    this.getProject();
    //this.getDefaultUsers();
    this.getAnexos();
    this.handler();

    jQuery("#btnConfirmarAdiar").hide();
    jQuery("#btnConfirmarCancelar").hide();
    jQuery("#btnConfirmarConcluir").hide();
    jQuery("#btnEditar").hide();




  }

  public render(): React.ReactElement<ISstDetalhesProps> {

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
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtName">Nome</label><span className="required"> *</span>
                      <br /><span id="txtName"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtCategoria">Categoria</label><span className="required"> *</span>
                      <br /><span id="txtCategoria">sdfsdffds</span>
                    </div>
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtTipo">Tipo</label><span className="required"> *</span>
                      <br /><span id="txtTipo"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtOwner">Owner</label><span className="required"> *</span>
                      <br /><span id="txtOwner"></span>

                    </div>
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtParticipantes">Participantes</label><span className="required"> *</span>
                      <br /><span id="txtParticipantes"></span>
                    </div>
                  </div>
                </div>


                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtDescricaoProduto">Descrição do Produto / Serviço</label><span className="required"> *</span>
                      <br /><span id="txtDescricaoProduto"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtRequisitosCriticos">Requisitos críticos</label><span className="required"> *</span>
                      <br /><span id="txtRequisitosCriticos"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtCliente">Cliente</label><span className="required"> *</span>
                      <br /><span id="txtCliente"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtOMPDocuments">Documentos OMP</label><span className="required"> *</span>
                      <br /><span id="txtOMPDocuments"></span>
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
                      {this.state.itemsListAnexosItem.map((item, key) => {

                        _pos++;
                        var txtAnexoItem = "anexoItem" + _pos;
                        var btnExcluirAnexoitem = "btnExcluirAnexoitem" + _pos;

                        //var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Anexos')/items('${_projectID}')/AttachmentFiles`;
                        var url = this.props.siteurl;

                        var caminho = `${url}/Lists/Projects/Attachments/${_projectID}/${item.FileName}`;

                        return (

                          <><a id={txtAnexoItem} target='_blank' data-interception="off" href={caminho} title="">{item.FileName}</a><br></br></>


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

                          <><a id={txtAnexoItem2} target='_blank' data-interception="off" href={caminho} title="">{item.Name}</a><br></br></>

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
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingAcoes" data-toggle="collapse" data-target="#collapseAcoes" aria-expanded="true" aria-controls="collapseAcoes">
              <h5 className="mb-0 text-info">
                Ações
              </h5>
            </div>
            <div id="collapseAcoes" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">
                <br></br><div className="text-right">
                  <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
                  <button style={{ "margin": "2px" }} id="btnEditar" className="btn btn-secondary">Editar</button>
                  <button style={{ "margin": "2px" }} id="btnConfirmarAdiar" className="btn btn-success">Adiar</button>
                  <button style={{ "margin": "2px" }} id="btnConfirmarCancelar" className="btn btn-success">Cancelar</button>
                  <button style={{ "margin": "2px" }} id="btnConfirmarConcluir" className="btn btn-success">Concluir</button>
                  <br></br><br></br>
                </div>

              </div>
            </div>
          </div>

        </div>
      </div>


        <div className="modal fade" id="modalConfirmarAdiar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
              </div>
              <div className="modal-body">
                Deseja realmente adiar a Solicitação?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnAdiar" type="button" className="btn btn-primary">Sim</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalConfirmarCancelar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
              </div>
              <div className="modal-body">
                Deseja realmente cancelar a Solicitação?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnCancelar" type="button" className="btn btn-primary">Sim</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalConfirmarConcluir" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
              </div>
              <div className="modal-body">
                Deseja realmente concluir a Solicitação?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnConcluir" type="button" className="btn btn-primary">Sim</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoAdiar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Solicitação adiada com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoAdiar" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoCancelar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Solicitação cancelada com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoCancelar" className="btn btn-primary">OK</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="modalSucessoConcluir" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Solicitação concluída com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucessoConcluir" className="btn btn-primary">OK</button>
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

    //Project_x0020_TXT eq _projectTitle
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

        console.log("resultData milestone", resultData);

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
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Projects List')/items?$select=ID,Title,ProjCategory,Project_x0020_type,AssignedTo/ID,AssignedTo/Title,Participants/ID,Participants/Title,Product_x0020_description_x0020_,Critical_x0020_requirements,Client/ID,Client/Title,OMP_x0020_documents,ProjStatus,SiteNovoSO,Owner_x0020_2,Participants_x0020_2&$expand=AssignedTo,Participants,Client&$filter=ID eq ` + _projectID,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        //  console.log("resultData doc", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var id = resultData.d.results[i].ID;
            var title = resultData.d.results[i].Title;

            _projectTitle = title;

            var status = resultData.d.results[i].ProjStatus;
            var nome = resultData.d.results[i].Title;
            var category = resultData.d.results[i].ProjCategory;
            var tipo = resultData.d.results[i].Project_x0020_type;
            var arrOwner = [];
            var arrParticipants = [];
            var descricaoProduto = resultData.d.results[i].Product_x0020_description_x0020_;
            var requisitosCriticos = resultData.d.results[i].Critical_x0020_requirements;
            var omp = resultData.d.results[i].OMP_x0020_documents;
            var arrCliente = [];
            var siteNovo = resultData.d.results[i].SiteNovoSO;
            _siteNovo = siteNovo;

            console.log("siteNovo", siteNovo);

            jQuery("#txtID").html(id);
            jQuery("#txtStatus").html(status);
            jQuery("#txtName").html(nome);
            jQuery("#txtCategoria").html(category);
            jQuery("#txtTipo").html(tipo);
            jQuery("#txtDescricaoProduto").html(descricaoProduto);
            jQuery("#txtRequisitosCriticos").html(requisitosCriticos);
            jQuery("#txtOMPDocuments").html(omp);

            if (siteNovo) {

              if (resultData.d.results[i].AssignedTo.hasOwnProperty('results')) {

                for (let x = 0; x < resultData.d.results[i].AssignedTo.results.length; x++) {

                  arrOwner.push(resultData.d.results[i].AssignedTo.results[x].Title);

                }

              }


              if (resultData.d.results[i].Participants.hasOwnProperty('results')) {

                for (let x = 0; x < resultData.d.results[i].Participants.results.length; x++) {

                  arrParticipants.push(resultData.d.results[i].Participants.results[x].Title);

                }

              }

            }

            else {

              arrOwner.push(resultData.d.results[i].Owner_x0020_2);
              arrParticipants.push(resultData.d.results[i].Participants_x0020_2);

            }


            var arrCliente = [];

            console.log("resultData.d.results[i].Client", resultData.d.results[i].Client);
            console.log("resultData.d.results[i].Client.length", resultData.d.results[i].Client.results.length);

            for (let x = 0; x < resultData.d.results[i].Client.results.length; x++) {

              console.log("resultData.d.results[i].Client.results[i].Title", resultData.d.results[i].Client.results[x].Title);

              arrCliente.push(resultData.d.results[i].Client.results[x].Title);

            }

            jQuery("#txtOwner").html(arrOwner.toString());
            jQuery("#txtParticipantes").html(arrParticipants.join(' - '));
            jQuery("#txtCliente").html(arrCliente.join(' - '));


            console.log("status", status);

            if ((status != "Concluída") && (status != "Cancelada")) {

              if (_grupos.indexOf("SST - Elaboradores") !== -1) {

                setTimeout(() => {

                  jQuery("#btnEditar").show();

                }, 1000);

              }

            }

            if (status == "Em Andamento") {

              if (arrOwner.indexOf(_correntUser) !== -1) {

                setTimeout(() => {

                  jQuery("#btnConfirmarAdiar").show();
                  jQuery("#btnConfirmarCancelar").show();
                  jQuery("#btnConfirmarConcluir").show();

                }, 1000);

              }

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

  protected confirmarEditar(opcao) {

    if (opcao == "Adiar") jQuery("#modalConfirmarAdiar").modal({ backdrop: 'static', keyboard: false });
    if (opcao == "Cancelar") jQuery("#modalConfirmarCancelar").modal({ backdrop: 'static', keyboard: false });
    if (opcao == "Concluir") jQuery("#modalConfirmarConcluir").modal({ backdrop: 'static', keyboard: false });

  }

  protected async editar(opcao) {

    jQuery("#btnAdiar").prop("disabled", true);
    jQuery("#btnCancelar").prop("disabled", true);
    jQuery("#btnConcluir").prop("disabled", true);

    var status;

    if (opcao == "Editar") {
      window.location.href = `Solicitacao-Editar.aspx?ProjectID=` + _projectID;
    }

    if (opcao == "Adiar") {

      status = "Adiada";

    }

    if (opcao == "Cancelar") {

      status = "Cancelada";

    }

    if (opcao == "Concluir") {

      status = "Concluída";

    }


    await _web.lists
      .getByTitle("Projects List")
      .items.getById(_projectID).update({
        ProjStatus: status
      })
      .then(response => {

        jQuery("#modalConfirmarAdiar").modal('hide');
        jQuery("#modalConfirmarCancelar").modal('hide');
        jQuery("#modalConfirmarConcluir").modal('hide');

        if (opcao == "Adiar") jQuery("#modalSucessoAdiar").modal({ backdrop: 'static', keyboard: false });
        if (opcao == "Cancelar") jQuery("#modalSucessoCancelar").modal({ backdrop: 'static', keyboard: false });
        if (opcao == "Concluir") jQuery("#modalSucessoConcluir").modal({ backdrop: 'static', keyboard: false });

      })
      .catch((error: any) => {
        console.log(error);
      })



  }

  protected voltar() {
    history.back();
  }

  protected async fecharSucessoEditar(opcao) {

    jQuery("#modalSucessoAdiar").modal('hide');
    jQuery("#modalSucessoCancelar").modal('hide');
    jQuery("#modalSucessoConcluir").modal('hide');

    window.location.href = `Solicitacao-Todos.aspx`;

  }

}
