import { SPHttpClient } from "@microsoft/sp-http";
export interface IReactWebpartDemoProps {
  description: string;
  spHttpClient: SPHttpClient;
  currentSiteUrl: string;
}
