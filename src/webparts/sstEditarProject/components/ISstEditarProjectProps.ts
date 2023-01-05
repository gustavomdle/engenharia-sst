import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISstEditarProjectProps {
  description: string;
  context: WebPartContext;
  siteurl: string;
}
