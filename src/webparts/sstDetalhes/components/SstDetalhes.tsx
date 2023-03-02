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
var _projectMilestone;
var _linhaAnexos = "";
var _strRequisitosCriticos = "";
var _strV3Comments = "";
var _strComment = "";

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
  itemsListForum: [
    {
      "ID": any,
      "Project": { "Title": any },
      "Title": any,
      "Created": any,
      "Body": any,
      "Author": { "Title": any },
      "To": any,
      "Folder": { "ItemCount": any },
    }
  ],
  itemsListForumRespostas: [
    {
      "ID": any,
      "Project": { "Title": any },
      "Title": any,
      "Created": any,
      "Body": any,
      "Author": any,
      "To": { "Title": any },
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
  itemsListTarefas: any[],
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
      itemsListForum: [
        {
          "ID": "",
          "Project": { "Title": "" },
          "Title": "",
          "Created": "",
          "Body": "",
          "Author": { "Title": "" },
          "To": "",
          "Folder": { "ItemCount": "" },
        }
      ],
      itemsListForumRespostas: [
        {
          "ID": "",
          "Project": { "Title": "" },
          "Title": "",
          "Created": "",
          "Body": "",
          "Author": "",
          "To": { "Title": "" },
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
      itemsListTarefas: [],
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

    // document
    //   .getElementById("btnConfirmarAdiar")
    //   .addEventListener("click", (e: Event) => this.confirmarEditar("Adiar"));

    // document
    //   .getElementById("btnConfirmarCancelar")
    //   .addEventListener("click", (e: Event) => this.confirmarEditar("Cancelar"));

    // document
    //   .getElementById("btnConfirmarConcluir")
    //   .addEventListener("click", (e: Event) => this.confirmarEditar("Concluir"));

    // document
    //   .getElementById("btnEditar")
    //   .addEventListener("click", (e: Event) => this.editar("Editar"));

    // document
    //   .getElementById("btnAdiar")
    //   .addEventListener("click", (e: Event) => this.editar("Adiar"));

    // document
    //   .getElementById("btnCancelar")
    //   .addEventListener("click", (e: Event) => this.editar("Cancelar"));

    // document
    //   .getElementById("btnConcluir")
    //   .addEventListener("click", (e: Event) => this.editar("Concluir"));

    // document
    //   .getElementById("btnSucessoAdiar")
    //   .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

    // document
    //   .getElementById("btnSucessoCancelar")
    //   .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

    // document
    //   .getElementById("btnSucessoConcluir")
    //   .addEventListener("click", (e: Event) => this.fecharSucessoEditar("Salvar"));

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
        text: "Title",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Priority",
        text: "Priority",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Status",
        text: "Issue Status",
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
        text: "Assigned To",
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
        text: "Due Date",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var data = new Date(row.DueDate);
          console.log("data issues", data)
          if (data != null) {
            var dtdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
            if (dtdata == "31/12/1969") dtdata = "";
          }
          else dtdata = "";
          return dtdata;
        }
      },
      {
        dataField: "",
        text: "Description",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var description = row.Comment;
          return <div dangerouslySetInnerHTML={{ __html: `${description}` }} />;

        }
      },
      {
        dataField: "",
        text: "Comments",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var idLista = this.props.idListaIssues;
          var id = row.ID;

          // console.log("_projectID issues", _projectID);

          var soapPack = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Body>
            <GetVersionCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">
              <strlistID>${idLista}</strlistID>
              <strlistItemID>${id}</strlistItemID>
              <strFieldName>V3Comments</strFieldName>
            </GetVersionCollection>
          </soap:Body>
        </soap:Envelope>`;


          console.log("soapPack issues", soapPack);

          $.ajax({
            type: "POST",
            url: this.props.siteurl + '/_vti_bin/lists.asmx',
            data: soapPack,
            dataType: "xml",
            async: false,
            contentType: "text/xml; charset=\"utf-8\"",
            success: function (xData1, status) {

              console.log("xData 1", xData1)

              $(xData1).find("Versions").find("Version").each(function () {

                var textoEditor2 = $(this).attr("Editor");

                console.log("textoEditor", textoEditor2);

                var editor1 = textoEditor2.substring(textoEditor2.indexOf("#") + 1);
                var editor2 = editor1.split('#')[0];

                var dtModified = new Date($(this).attr("Modified"));
                //  dtModified = moment(dtModified).format('DD/MM/YYYY HH:mm');

                var dtModified = new Date($(this).attr("Modified"));
                var formDtdata = ("0" + dtModified.getDate()).slice(-2) + '/' + ("0" + (dtModified.getMonth() + 1)).slice(-2) + '/' + dtModified.getFullYear() + ' ' + ("0" + (dtModified.getHours())).slice(-2) + ':' + ("0" + (dtModified.getMinutes())).slice(-2);


                _strV3Comments += "<span style='color:#004b87'>" + editor2 + "(" + formDtdata + ")</span><br />" + $(this).attr("V3Comments");
                //strRequisitosCriticos = strRequisitosCriticos.replace("undefined", "");
                //strRequisitosCriticos = strRequisitosCriticos.replace(",(", " (");
                //strRequisitosCriticos = strRequisitosCriticos.replace(",,", ",");

              });

              //console.log("strProdutoDescricao",strProdutoDescricao);
              //jQuery("#txtRequisitosCriticos").html(strRequisitosCriticos);
            },
            error: function (e) {
              console.log("e", e);
            }
          });


          return <div dangerouslySetInnerHTML={{ __html: `${_strV3Comments}` }} />;

        }
      },
      {
        dataField: "",
        text: "Anexos",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var id = row.ID;

          var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Project Issues')/items('${id}')/AttachmentFiles`;

          //console.log("url anexo", url);

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

                // console.log("resultData anexos RelatedIssues", resultData);


                if (resultData.d.results.length > 0) {

                  for (var i = 0; i < resultData.d.results.length; i++) {

                    var caminho = encodeURI(resultData.d.results[i].ServerRelativeUrl);

                    //   console.log("caminho arquivo", caminho);

                    _linhaAnexos += `<a target='_blank' data-interception="off" href=${caminho} >${resultData.d.results[i].FileName}</a><br></br>`;

                  }

                }

              },
              error: function (xhr, status, error) {
                console.log("Falha anexo");
              }
            })

          return <div dangerouslySetInnerHTML={{ __html: `${_linhaAnexos}` }} />;

        }
      },

    ]

    const tablecolumnsRelatedMilestones = [
      {
        dataField: "Title",
        text: "Milestone",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Complete",
        text: "Complete",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var data = row.Complete;
          var valor = "No";
          if (data == false) valor = "Yes";
          return valor;
        }
      },
      {
        dataField: "DueDate",
        text: "Due Date",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var data = new Date(row.DueDate);
          console.log("data", data);
          if (row.DueDate != null) {
            var dtdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
          }
          else dtdata = "";
          return dtdata;
        }
      },

      // {
      //   dataField: "ProjComments",
      //   text: "Comentários",
      //   headerStyle: { "backgroundColor": "#bee5eb" },
      //   classes: 'headerPreStage',
      //   headerClasses: 'text-center',
      //   formatter: (rowContent, row) => {

      //     var comentarios = row.ProjComments;
      //     var vlrComentario = "";

      //     if (comentarios != null) vlrComentario = row.ProjComments;

      //     return <div dangerouslySetInnerHTML={{ __html: `${vlrComentario}` }} />;

      //   }
      // },

      {
        dataField: "",
        text: "",
        headerStyle: { "backgroundColor": "#bee5eb", "width": "80px" },
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var id = row.ID;

          return (

            <button onClick={async () => { this.abrirModalRelatedMilestones(id); }} className="btn btn-info btnCustom btn-sm">Detalhes</button>

          )

        }
      }

    ]


    const tablecolumnsTarefas = [
      {
        dataField: "Title",
        text: "Title",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Milestone",
        text: "Milestone",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Priority",
        text: "Priority",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "CostDays",
        text: "Qtd in Hours",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
      },
      {
        dataField: "Status",
        text: "Status",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
      },
      {
        dataField: "PercentComplete",
        text: "% Complete",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
      },
      {
        dataField: "AtribuidoA",
        text: "Assigned To",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
      },
      {
        dataField: "Body",
        text: "Description",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          return <div dangerouslySetInnerHTML={{ __html: `${row.Body}` }} />;

        }
      },
      {
        dataField: "StartDate",
        text: "Start Date",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var data = new Date(row.StartDate);
          console.log("data", data);
          if (row.Created != null) {
            var dtdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
          }
          else dtdata = "";
          return dtdata;
        }
      },
      {
        dataField: "DueDate",
        text: "Due Date",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage text-center',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {
          var data = new Date(row.DueDate);
          console.log("data", data);
          if (row.DueDate != null) {
            var dtdata = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
          }
          else dtdata = "";
          return dtdata;
        }
      },
      {
        dataField: "",
        text: "Anexos",
        headerStyle: { "backgroundColor": "#bee5eb" },
        classes: 'headerPreStage',
        headerClasses: 'text-center',
        formatter: (rowContent, row) => {

          var id = row.ID;

          var url = `${this.props.siteurl}/_api/web/lists/getByTitle('Project Tasks')/items('${id}')/AttachmentFiles`;

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

                console.log("resultData anexos tasks", resultData);


                if (resultData.d.results.length > 0) {

                  for (var i = 0; i < resultData.d.results.length; i++) {

                    var caminho = encodeURI(resultData.d.results[i].ServerRelativeUrl);

                    console.log("caminho arquivo", caminho);

                    _linhaAnexos += `<a target='_blank' data-interception="off" href=${caminho} >${resultData.d.results[i].FileName}</a><br></br>`;

                  }

                }

              },
              error: function (xhr, status, error) {
                console.log("Falha anexo");
              }
            })

          return <div dangerouslySetInnerHTML={{ __html: `${_linhaAnexos}` }} />;

        }
      },



    ]



    return (


      <><div id="container">
        <div id="accordion">

          <div className="card">
            <div className="card-header btn" id="headingProjectInformation" data-toggle="collapse" data-target="#collapseProjectInformation" aria-expanded="true" aria-controls="collapseProjectInformation">
              <h5 className="mb-0 text-info">
                Project Information
              </h5>
            </div>
            <div id="collapseProjectInformation" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtName">Project Name</label><span className="required"> *</span>
                      <br /><span id="txtName"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtStatus">Status</label><span className="required"> *</span>
                      <br /><span id="txtStatus"></span>
                    </div>
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtCategoria">Category</label><span className="required"> *</span>
                      <br /><span id="txtCategoria"></span>
                    </div>

                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtTipo">Project type</label><span className="required"> *</span>
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
                      <label className="text-info" htmlFor="txtParticipantes">Participants</label><span className="required"> *</span>
                      <br /><span id="txtParticipantes"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtOwner">Start Date</label><span className="required"> *</span>
                      <br /><span id="txtDataInicial"></span>

                    </div>
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtDataFinal">End Date</label><span className="required"> *</span>
                      <br /><span id="txtDataFinal"></span>
                    </div>
                  </div>
                </div>


                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtDescricaoProduto">Product description / Service</label><span className="required"> *</span>
                      <br /><span id="txtDescricaoProduto"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtRequisitosCriticos">Critical requirements</label><span className="required"> *</span>
                      <br /><span id="txtRequisitosCriticos"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtCliente">Client</label><span className="required"> *</span>
                      <br /><span id="txtCliente"></span>
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group border m-1 col-md">
                      <label className="text-info" htmlFor="txtOMPDocuments">OMP documents</label><span className="required"> *</span>
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
                Attachment
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
                  <BootstrapTable bootstrap4 striped responsive condensed hover={false} className="gridTodosItens" id="gridTodosItensRelatedMilestones" keyField='id' data={this.state.itemsListRelatedMilestones} columns={tablecolumnsRelatedMilestones} headerClasses="header-class" />
                </div>
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header btn" id="headingForum" data-toggle="collapse" data-target="#collapseForum" aria-expanded="true" aria-controls="collapseForum">
              <h5 className="mb-0 text-info">
                Forum
              </h5>
            </div>
            <div id="collapseForum" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">
                <div id='conteudoForum'>

                  {this.state.itemsListForum.map((item, key) => {

                    var id = item.ID;
                    var arrTo = [];
                    var countRespostas = 0;

                    var criadoPor = item.Author.Title;
                    var to = item.To;

                    console.log("to", to);

                    if (to != "") {

                      let naoContemValor = to.hasOwnProperty('__deferred');

                      console.log("naoContemValor", naoContemValor);

                      if (!naoContemValor) {

                        for (var i = 0; i < to.results.length; i++) {

                          arrTo.push(to.results[i].Title);

                        }

                      }

                    }
                    // 

                    var criado = new Date(item.Created);

                    var dtcriado = ("0" + criado.getDate()).slice(-2) + '/' + ("0" + (criado.getMonth() + 1)).slice(-2) + '/' + criado.getFullYear() + ' ' + ("0" + (criado.getHours())).slice(-2) + ':' + ("0" + (criado.getMinutes())).slice(-2);
                    // var respostas = item.Respostas;
                    var respostas = item.Folder.ItemCount;
                    console.log("respostas", respostas);
                    var assunto = item.Title;

                    var corpo = item.Body;

                    //respostas = "1";

                    if (respostas != 0) {

                      return (

                        <div className="p-0 mb-0 bg-light text-dark rounded comment ${area2}">

                          <div className="p-3 mb-2 alert-danger text-dark rounded-top ">

                            <b>Comentário postado por:</b> {criadoPor} em {dtcriado}<br></br>
                            <b>Para:</b> {arrTo.toString()}<br></br>
                            <b>Respostas:</b> {respostas}


                          </div>
                          <br />
                          <div className="p-3">

                            <h4>{assunto}</h4>

                            <div dangerouslySetInnerHTML={{ __html: `${corpo}` }} />

                            <br /><br />
                            <button id='btnRespostas' onClick={async () => { this.abrirModalRespostas(id); }} type="button" className="btn btn-info">Respostas</button>
                            <br /><br />

                          </div>

                        </div>

                      );

                    } else {

                      return (

                        <div className="p-0 mb-0 bg-light text-dark rounded comment ${area2}">

                          <div className="p-3 mb-2 alert-danger text-dark rounded-top ">

                            <b>Comentário postado por:</b> {criadoPor} em {dtcriado}<br></br>
                            <b>Para:</b> {arrTo.toString()}<br></br>
                            <b>Respostas:</b> {respostas}

                          </div>
                          <br />
                          <div className="p-3">

                            <h4>{assunto}</h4>

                            <div dangerouslySetInnerHTML={{ __html: `${corpo}` }} />

                          </div>

                        </div>

                      );

                    }



                  })}



                </div>
              </div>
            </div>
          </div>

          <br></br><div className="text-right">
            <button style={{ "margin": "2px" }} type="submit" id="btnVoltar" className="btn btn-secondary">Voltar</button>
            <br></br><br></br>
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

        <div className="modal fade" id="modalDetalhesMilestones" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog modalLargura1100" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Related Milestones</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">

                <div className="form-row">
                  <div className="form-group col-md border m-1">
                    <label className="text-info" htmlFor="txtProjectMilestone">Project Milestone</label><br></br>
                    <span className='labelDetalhes' id='txtProjectMilestone'></span>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md border m-1">
                    <label className="text-info" htmlFor="txtProject">Project</label><br></br>
                    <span className='labelDetalhes' id='txtProject'></span>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md border m-1">
                    <label className="text-info" htmlFor="txtDueDate">Due Date</label><br></br>
                    <span className='labelDetalhes' id='txtDueDate'></span>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md border m-1">
                    <label className="text-info" htmlFor="txtComplete">Complete</label><br></br>
                    <span className='labelDetalhes' id='txtComplete'></span>
                  </div>
                </div>

                <div className="form-row">
                  <div className="form-group col-md border m-1">
                    <label className="text-info" htmlFor="txtComments">Comments</label><br></br>
                    <span className='labelDetalhes' id='txtComments'></span>
                  </div>
                </div>

                <br></br>

                <BootstrapTable bootstrap4 striped responsive condensed hover={false} className="gridTodosItens" id="gridTodosItensTarefas" keyField='id' data={this.state.itemsListTarefas} columns={tablecolumnsTarefas} headerClasses="header-class" />

              </div>
              <div className="modal-footer">
              </div>
            </div>
          </div>
        </div>


        <div className="modal fade" id="modalRespostas" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog modalLargura1100" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Respostas</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">

                {this.state.itemsListForumRespostas.map((item, key) => {

                  console.log("item", item);

                  var criadoPor = "";

                  if (item.Author != "") {

                    var arrAuthor = item.Author;

                    console.log("arrAuthor", arrAuthor[0].title);
                    criadoPor = arrAuthor[0].title;

                  }

                  console.log("criadoPor", criadoPor);

                  //  console.log("item.Created2010",item.Created2010);
                  // var criado = new Date(item.Created2010);
                  // console.log("criado",criado);
                  // var dtcriado = ("0" + criado.getDate()).slice(-2) + '/' + ("0" + (criado.getMonth() + 1)).slice(-2) + '/' + criado.getFullYear() + ' ' + ("0" + (criado.getHours())).slice(-2) + ':' + ("0" + (criado.getMinutes())).slice(-2);
                  var assunto = item.Title;

                  var corpo = item.Body;

                  return (

                    <div className="p-0 mb-0 bg-light text-dark rounded comment ${area2}">

                      <div className="p-3 mb-2 alert-danger text-dark rounded-top ">

                        <b>Comentário postado por:</b> {criadoPor} em {item.Created}<br></br>

                      </div>
                      <br />
                      <div className="p-3">

                        <h4>{assunto}</h4>

                        <div dangerouslySetInnerHTML={{ __html: `${corpo}` }} />

                      </div>

                    </div>

                  );

                })}

              </div>
              <div className="modal-footer">
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

    console.log("_projectTitle", _projectTitle);
    console.log("_siteNovo", _siteNovo);

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


    var reactForum = this;

    console.log("_projectTitle", _projectTitle);

    jquery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/GetByTitle('BK Forum 3')/items?$top=50&$orderby= Created asc&$select=ID,Title,Project/Title,Created,Body,Author/Title,To/Title,Folder/ItemCount&$expand=Project,Author,To,Folder&$filter=Project/Title eq '${_projectTitle}' `,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (resultData) {

        console.log("resultData Forum", resultData);

        if (resultData.d.results.length > 0) {
          jQuery("#tabelaPreStageSoftware").show();
          reactForum.setState({
            itemsListForum: resultData.d.results
          });
        } else {
          jQuery("#conteudoForum").hide();
        }


      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


  }


  protected getProject() {

    jQuery.ajax({
      url: `${this.props.siteurl}/_api/web/lists/getbytitle('Projects List')/items?$select=ID,Title,ProjCategory,Project_x0020_type,AssignedTo/ID,AssignedTo/Title,Participants/ID,Participants/Title,Product_x0020_description_x0020_,Critical_x0020_requirements,Client/ID,Client/Title,OMP_x0020_documents,ProjStatus,SiteNovoSO,Owner_x0020_2,Participants_x0020_2,StartDate,EndDate,Status_x0020_Projeto&$expand=AssignedTo,Participants,Client&$filter=ID eq ` + _projectID,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: async (resultData) => {

        console.log("resultData doc", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var id = resultData.d.results[i].ID;
            var title = resultData.d.results[i].Title;

            _projectTitle = title;

            var status = resultData.d.results[i].Status_x0020_Projeto;
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

            var dataInicial = new Date(resultData.d.results[i].StartDate);

            if (resultData.d.results[i].StartDate != null) {

              var dtdataInicial = ("0" + dataInicial.getDate()).slice(-2) + '/' + ("0" + (dataInicial.getMonth() + 1)).slice(-2) + '/' + dataInicial.getFullYear();

            }
            else {

              var dtdataInicial = "";
            }

            var dataFinal = new Date(resultData.d.results[i].EndDate);

            if (resultData.d.results[i].EndDate != null) {

              var dtdataFinal = ("0" + dataFinal.getDate()).slice(-2) + '/' + ("0" + (dataFinal.getMonth() + 1)).slice(-2) + '/' + dataFinal.getFullYear();

            } else {

              var dtdataFinal = "";
            }

            jQuery("#txtID").html(id);
            jQuery("#txtStatus").html(status);
            jQuery("#txtName").html(nome);
            jQuery("#txtCategoria").html(category);
            jQuery("#txtTipo").html(tipo);
            // jQuery("#txtDescricaoProduto").html(descricaoProduto);
            //jQuery("#txtRequisitosCriticos").html(requisitosCriticos);
            jQuery("#txtOMPDocuments").html(omp);
            jQuery("#txtDataInicial").html(dtdataInicial);
            jQuery("#txtDataFinal").html(dtdataFinal);

            if (siteNovo) {

              if (resultData.d.results[i].AssignedTo.hasOwnProperty('results')) {

                for (let x = 0; x < resultData.d.results[i].AssignedTo.results.length; x++) {

                  arrOwner.push(resultData.d.results[i].AssignedTo.results[x].Title);

                }

              }




            }

            else {

              arrOwner.push(resultData.d.results[i].Owner_x0020_2);
              // arrParticipants.push(resultData.d.results[i].Participants);

            }

            if (resultData.d.results[i].Participants.hasOwnProperty('results')) {

              for (let x = 0; x < resultData.d.results[i].Participants.results.length; x++) {

                arrParticipants.push(resultData.d.results[i].Participants.results[x].Title);

              }

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

            // if ((status != "Concluída") && (status != "Cancelada")) {

            //   if (_grupos.indexOf("SST - Elaboradores") !== -1) {

            //     setTimeout(() => {

            //       jQuery("#btnEditar").show();

            //     }, 1000);

            //   }

            // }

            // if (status == "Em Andamento") {

            //   if (arrOwner.indexOf(_correntUser) !== -1) {

            //     setTimeout(() => {

            //       jQuery("#btnConfirmarAdiar").show();
            //       jQuery("#btnConfirmarCancelar").show();
            //       jQuery("#btnConfirmarConcluir").show();

            //     }, 1000);

            //   }

            // }





          }
        }




      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }

    })

    var idLista = this.props.idListaProject;

    if (idLista == "") {

      alert("GUID da lista não encontrado nas configuraçãoes da webpart");

    }

    var soapPack = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                    <soap:Body>
                      <GetVersionCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <strlistID>${idLista}</strlistID>
                        <strlistItemID>${_projectID}</strlistItemID>
                        <strFieldName>Product_x0020_description_x0020_</strFieldName>
                      </GetVersionCollection>
                    </soap:Body>
                  </soap:Envelope>`;

    $.ajax({
      type: "POST",
      url: this.props.siteurl + '/_vti_bin/lists.asmx',
      data: soapPack,
      dataType: "xml",
      async: false,
      contentType: "text/xml; charset=\"utf-8\"",
      success: function (xData, status) {

        var strDescricaoProduto = "";

        console.log("xData 1", xData)

        $(xData).find("Versions").find("Version").each(function () {

          var textoEditor = $(this).attr("Editor");

          console.log("textoEditor", textoEditor);

          var editor1 = textoEditor.substring(textoEditor.indexOf("#") + 1);
          var editor2 = editor1.split('#')[0];

          var dtModified = new Date($(this).attr("Modified"));
          //  dtModified = moment(dtModified).format('DD/MM/YYYY HH:mm');

          var dtModified = new Date($(this).attr("Modified"));
          var formDtdata = ("0" + dtModified.getDate()).slice(-2) + '/' + ("0" + (dtModified.getMonth() + 1)).slice(-2) + '/' + dtModified.getFullYear() + ' ' + ("0" + (dtModified.getHours())).slice(-2) + ':' + ("0" + (dtModified.getMinutes())).slice(-2);


          strDescricaoProduto += "<span style='color:#004b87'>" + editor2 + "(" + formDtdata + ")</span><br />" + $(this).attr("Product_x0020_description_x0020_");
          //  strDescricaoProduto = strDescricaoProduto.replace("undefined", "");
          //   strDescricaoProduto = strDescricaoProduto.replace(",(", " (");
          //  strDescricaoProduto = strDescricaoProduto.replace(",,", ",");

        });

        //console.log("strProdutoDescricao",strProdutoDescricao);
        jQuery("#txtDescricaoProduto").html(strDescricaoProduto);
      },
      error: function (e) {
        console.log("e", e);
      }
    });


    var soapPack2 = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                    <soap:Body>
                      <GetVersionCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                        <strlistID>${idLista}</strlistID>
                        <strlistItemID>${_projectID}</strlistItemID>
                        <strFieldName>Critical_x0020_requirements</strFieldName>
                      </GetVersionCollection>
                    </soap:Body>
                  </soap:Envelope>`;

    $.ajax({
      type: "POST",
      url: this.props.siteurl + '/_vti_bin/lists.asmx',
      data: soapPack2,
      dataType: "xml",
      async: false,
      contentType: "text/xml; charset=\"utf-8\"",
      success: function (xData, status) {

        var strRequisitosCriticos = "";

        console.log("xData 2", xData)

        $(xData).find("Versions").find("Version").each(function () {

          var textoEditor2 = $(this).attr("Editor");

          console.log("textoEditor", textoEditor2);

          var editor1 = textoEditor2.substring(textoEditor2.indexOf("#") + 1);
          var editor2 = editor1.split('#')[0];

          var dtModified = new Date($(this).attr("Modified"));
          //  dtModified = moment(dtModified).format('DD/MM/YYYY HH:mm');

          var dtModified = new Date($(this).attr("Modified"));
          var formDtdata = ("0" + dtModified.getDate()).slice(-2) + '/' + ("0" + (dtModified.getMonth() + 1)).slice(-2) + '/' + dtModified.getFullYear() + ' ' + ("0" + (dtModified.getHours())).slice(-2) + ':' + ("0" + (dtModified.getMinutes())).slice(-2);


          strRequisitosCriticos += "<span style='color:#004b87'>" + editor2 + "(" + formDtdata + ")</span><br />" + $(this).attr("Critical_x0020_requirements");
          //strRequisitosCriticos = strRequisitosCriticos.replace("undefined", "");
          //strRequisitosCriticos = strRequisitosCriticos.replace(",(", " (");
          //strRequisitosCriticos = strRequisitosCriticos.replace(",,", ",");

        });

        //console.log("strProdutoDescricao",strProdutoDescricao);
        jQuery("#txtRequisitosCriticos").html(strRequisitosCriticos);
      },
      error: function (e) {
        console.log("e", e);
      }
    });



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

  protected abrirModalRelatedMilestones(id) {


    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Milestones')/items?$top=50&$orderby= Created asc&$select=ID,Title,Complete,DueDate,ProjComments,Project_x0020_TXT&$filter=ID eq ` + id;

    jQuery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: function (resultData) {

        console.log("resultData milestone", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var projectMilestone = resultData.d.results[i].Title;
            _projectMilestone = projectMilestone
            var project = resultData.d.results[i].Project_x0020_TXT;
            var dueDate = resultData.d.results[i].DueDate;
            var complete = resultData.d.results[i].Complete;
            var comments = resultData.d.results[i].ProjComments;


            if (dueDate != null) {
              var data = new Date(dueDate);
              var resData = ("0" + data.getDate()).slice(-2) + '/' + ("0" + (data.getMonth() + 1)).slice(-2) + '/' + data.getFullYear();
            }
            else resData = "";

            var resComplete = "No";

            if (complete != null) {
              if (complete != false) {
                resComplete = complete;
              }
              else {
                resComplete = "No";
              }
            }


            jQuery("#txtProjectMilestone").html(projectMilestone);
            jQuery("#txtProject").html(project);
            jQuery("#txtDueDate").html(resData);
            jQuery("#txtComplete").html(resComplete);
            jQuery("#txtComments").html(comments);


          }


        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    var reactHandlerTarefas = this;

    console.log("_projectTitle", _projectTitle);
    console.log("_siteNovo", _siteNovo);

    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Tasks')/items?$top=50&$orderby= Created asc&$select=ID,Title,Priority,DueDate,Status,Milestone,AtribuidoA,CostDays,PercentComplete,Body,Created,StartDate&$filter=Milestone eq '${_projectMilestone}' `;
    // var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Project Tasks')/items?$top=50&$orderby= Created asc&$select=ID,Title&$filter=Milestone eq '${_projectMilestone}' `;

    jQuery.ajax({
      url: url,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      async: false,
      success: function (resultData) {

        console.log("resultData tarefas", resultData);

        if (resultData.d.results.length > 0) {
          jQuery("#tabelaPreStageSoftware").show();
          reactHandlerTarefas.setState({
            itemsListTarefas: resultData.d.results
          });
        }
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    setTimeout(async () => {
      jQuery("#modalDetalhesMilestones").modal({ backdrop: 'static', keyboard: false });
    }, 1000);




  }


  protected async abrirModalRespostas(id) {

    console.log("id", id);

    var reactHandlerTarefas = this;

    console.log("_projectTitle", _projectTitle);
    console.log("_siteNovo", _siteNovo);

    var url = `${this.props.siteurl}/_api/web/lists/getbytitle('Lista Base Forum 2')/items?$top=1&$orderby= Created asc&$select=ID,Body&$filter=Title eq '${id}' `;

    console.log("url abrirModalRespostas", url);

    jQuery.ajax({
      url: url,
      type: "GET",
      async: false,
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: async (resultData) => {

        console.log("resultData.d.results.length abrirModalRespostas", resultData);

        if (resultData.d.results.length > 0) {

          for (var i = 0; i < resultData.d.results.length; i++) {

            var title = resultData.d.results[i].Body;

            var listUri = `${this.props.context.pageContext.web.serverRelativeUrl}/Lists/BK Forum 3`;

            console.log("listUri", listUri);

            await _web.getList(listUri)
              .renderListDataAsStream({
                // ViewXml: '',
                FolderServerRelativeUrl: `${listUri}/${title}`
              })
              .then(r => {

                console.log("r1", r);

                var reactHandlerForumRespostas = this;


                setTimeout(() => {

                  reactHandlerForumRespostas.setState({
                    itemsListForumRespostas: r.Row
                  });

                }, 0);


              }).catch((error: any) => {
                console.log("Erro respostas: ", error);
              });



          }

        }



      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR.responseText);
      }
    });


    jQuery("#modalRespostas").modal({ backdrop: 'static', keyboard: false });



  }

}
