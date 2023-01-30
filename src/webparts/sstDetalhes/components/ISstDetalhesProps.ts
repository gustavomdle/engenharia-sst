import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISstDetalhesProps {
  description: string;
  context: WebPartContext;
  siteurl: string;
  idListaProject: string;
}
