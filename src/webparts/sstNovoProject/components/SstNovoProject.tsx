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

export interface IReactGetItemsState {

  itemsCategoria: [],
  itemsTipo: [],

}

export default class SstNovoProject extends React.Component<ISstNovoProjectProps, IReactGetItemsState> {

  public constructor(props: ISstNovoProjectProps, state: IReactGetItemsState) {
    super(props);
    this.state = {

      itemsCategoria: [],
      itemsTipo: [],
    };
  }

  public componentDidMount() {

    _web = new Web(this.props.context.pageContext.web.absoluteUrl);
    _caminho = this.props.context.pageContext.web.serverRelativeUrl;

    this.handler();


  }


  public render(): React.ReactElement<ISstNovoProjectProps> {
    return (


      <><div id="container">
        <div id="accordion">
          <div className="card">
            <div className="card-header btn" id="headingInformacoesProduto" data-toggle="collapse" data-target="#collapseInformacoesProduto" aria-expanded="true" aria-controls="collapseInformacoesProduto">
              <h5 className="mb-0 text-info">
              Project Information
              </h5>
            </div>
            <div id="collapseInformacoesProduto" className="collapse show" aria-labelledby="headingOne">
              <div className="card-body">

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md-6">
                      <label htmlFor="txtNomeProduto">Name</label><span className="required"> *</span>
                      <input type="text" className="form-control" id="txtNomeProduto" />
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="txtNomeProduto">Category</label><span className="required"> *</span>
                      <select id="ddlSistemaOperacional" className="form-control">
                        <option value="0" selected>Selecione...</option>
                        {this.state.itemsCategoria.map(function (item, key) {
                          return (
                            <option value={item}>{item}</option>
                          );
                        })}
                      </select>
                    </div>
                    <div className="form-group col-md-3">
                      <label htmlFor="txtNomeProduto">Type</label><span className="required"> *</span>
                      <select id="ddlSistemaOperacional" className="form-control">
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
                      <label htmlFor="txtNomeProduto">Owner</label><span className="required"> *</span>
                      <PeoplePicker
                        context={this.props.context as any}
                        //titleText="Aprovador Engenharia"
                        personSelectionLimit={1}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        onChange={this._getPeoplePickerItems.bind(this)}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        //className={ styles.label  }
                        ensureUser={true} />

                    </div>
                    <div className="form-group col-md">
                      <label htmlFor="txtNomeProduto">Participants</label><span className="required"> *</span>
                      <PeoplePicker
                        context={this.props.context as any}
                        //titleText="Aprovador Engenharia"
                        personSelectionLimit={20}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        onChange={this._getPeoplePickerItems.bind(this)}
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
                      <label htmlFor="txtNomeProduto">Product description / Service</label><span className="required"> *</span>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeProductDescription(text)} />
                    </div>
                  </div>
                </div>

                <div className="form-group">
                  <div className="form-row">
                    <div className="form-group col-md">
                      <label htmlFor="txtNomeProduto">Critical requirements</label><span className="required"> *</span>
                      <RichText className="editorRichTex" value=""
                        onChange={(text) => this.onTextChangeCriticalRequirements(text)} />
                    </div>
                  </div>
                </div>


              </div>
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


  }

  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }

  private onTextChangeProductDescription = (newText: string) => {
    //_outrasInformacoes = newText;
    return newText;
  }

  private onTextChangeCriticalRequirements = (newText: string) => {
    //_outrasInformacoes = newText;
    return newText;
  }


  

}
