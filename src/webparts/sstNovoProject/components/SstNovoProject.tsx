import * as React from 'react';
import styles from './SstNovoProject.module.scss';
import { ISstNovoProjectProps } from './ISstNovoProjectProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

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
import { PrimaryButton, Stack, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { SiteUser } from 'sp-pnp-js/lib/sharepoint/siteusers';

require("../../../../node_modules/bootstrap/dist/css/bootstrap.min.css");
require("../../../../css/estilos.css");

var _web;
var _caminho;
var _arrOwner = [];
var _arrParticipants = [];
var _productDescription;
var _criticalRequirements;
var _projectID;
var _size: number = 0;


export interface IReactGetItemsState {

  itemsCliente: [
    {
      "ID": "",
      "Title": "",
    }],
  itemsCategoria: [],
  itemsTipo: [],

}

export default class SstNovoProject extends React.Component<ISstNovoProjectProps, IReactGetItemsState> {

  public constructor(props: ISstNovoProjectProps, state: IReactGetItemsState) {
    super(props);
    this.state = {

      itemsCliente: [
        {
          "ID": "",
          "Title": "",
        }],
      itemsCategoria: [],
      itemsTipo: [],
    };
  }

  public componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);
    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    document
      .getElementById("btnConfirmarSalvar")
      .addEventListener("click", (e: Event) => this.validar());

    document
      .getElementById("btnSalvar")
      .addEventListener("click", (e: Event) => this.salvar());

    document
      .getElementById("btnSucesso")
      .addEventListener("click", (e: Event) => this.fecharSucesso());

    jQuery("#conteudoLoading").html(`<br/><br/><img style="height: 80px; width: 80px" src='${_caminho}/SiteAssets/loading.gif'/>
      <br/>Aguarde....<br/><br/>
      Dependendo do tamanho do anexo e a velocidade<br>
       da Internet essa ação pode demorar um pouco. <br>
       Não fechar a janela!<br/><br/>`);

    this.handler();


  }


  public render(): React.ReactElement<ISstNovoProjectProps> {
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
                      <select id="ddlCategory" className="form-control">
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
                      <select id="ddlType" className="form-control">
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
                        //className={ styles.label  }
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
                        //className={ styles.label  }
                        ensureUser={true} />
                    </div>
                  </div>
                </div>


                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtProduct">Descrição do Produto / Serviço</label><span className="required"> *</span>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeProductDescription(text)} />
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtCritical">Requisitos críticos</label><span className="required"> *</span>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeCriticalRequirements(text)} />
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
                                  return (
                                    <option className="optArea" value={item.ID}>{item.Title}</option>
                                  );

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

              </div>
            </div>
          </div>




        </div>
      </div>

        <br></br><div className="text-right">
          <button id="btnConfirmarSalvar" className="btn btn-success">Salvar</button>
        </div>

        <div className="modal fade" id="modalConfirmar" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Confirmação</h5>
              </div>
              <div className="modal-body">
                Deseja realmente criar uma novo Solicitação?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancelar</button>
                <button id="btnSalvar" type="button" className="btn btn-primary">Sim</button>
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

        <div className="modal fade" id="modalSucesso" tabIndex={-1} role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="exampleModalLabel">Alerta</h5>
              </div>
              <div className="modal-body">
                Solicitação criada com sucesso!
              </div>
              <div className="modal-footer">
                <button type="button" id="btnSucesso" className="btn btn-primary">OK</button>
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



  }


  protected validar() {

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

    jQuery("#modalConfirmar").modal({ backdrop: 'static', keyboard: false });

  }


  protected async salvar() {

    jQuery("#modalConfirmar").modal('hide');
    jQuery("#modalCarregando").modal({ backdrop: 'static', keyboard: false });

    var name = jQuery("#txtName").val();
    var category = jQuery("#ddlCategory option:selected").val();
    var type = jQuery("#ddlType option:selected").val();
    var ompDocuments = jQuery("#txtOMPDocuments").val();

    var arrPeoplepickerOwner = [];
    for (let i = 0; i < _arrOwner.length; i++) {
      arrPeoplepickerOwner.push(_arrOwner[i]["id"]);
    }

    var arrPeoplepickerParticipants = [];
    for (let i = 0; i < _arrParticipants.length; i++) {
      arrPeoplepickerParticipants.push(_arrParticipants[i]["id"]);
    }

    var arrClient = Array.prototype.slice.call(document.querySelectorAll('#ddlCliente2 option'), 0).map(function (v, i, a) {
      return v.value;
    });

    console.log("arrPeoplepickerOwner", arrPeoplepickerOwner);

    await _web.lists
      .getByTitle("Projects List")
      .items.add({
        Title: name,
        ProjCategory: category,
        Project_x0020_type: type,
        AssignedToId: { 'results': arrPeoplepickerOwner },
        ParticipantsId: { 'results': arrPeoplepickerParticipants },
        Product_x0020_description_x0020_: _productDescription,
        Critical_x0020_requirements: _criticalRequirements,
        ClientId: { "results": arrClient },
        OMP_x0020_documents: ompDocuments
      })
      .then(response => {
        _projectID = response.data.ID;
        console.log("criou!!");
        this.upload();
      })
      .catch((error: any) => {
        console.log(error);
      })

  }


  protected upload() {

    console.log("Entrou no upload");

    var files = (document.querySelector("#input") as HTMLInputElement).files;
    var file = files[0];

    //console.log("files.length", files.length);

    if (files.length != 0) {

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
                    jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false })
                  }
                });
            }


          }).catch(err => {
            console.log("err", err);
          });



      }).catch(err => {
        console.log("err", err);
      });

      //const folderAddResult = _web.folders.add(`${_caminho}/Anexos/${_idProposta}`);
      //console.log("foi");

    } else {

      _web.lists.getByTitle("Imagens").rootFolder.folders.add(`${_projectID}`).then(data => {

        console.log("Gravou!!");
        jQuery("#conteudoLoading").modal('hide');
        jQuery("#modalSucesso").modal({ backdrop: 'static', keyboard: false });

      }).catch(err => {
        console.log("err", err);
      });

    }



  }


  protected async fecharSucesso() {

    jQuery("#modalSucesso").modal('hide');
    window.location.href = `Solicitacao-Editar.aspx?ProjectID=` + _projectID;

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




}
