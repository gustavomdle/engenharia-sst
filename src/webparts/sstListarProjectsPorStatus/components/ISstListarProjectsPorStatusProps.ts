import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISstListarProjectsPorStatusProps {
  description: string;
  context: WebPartContext;
  siteurl: string;
  statusSolicitacao: string
}
