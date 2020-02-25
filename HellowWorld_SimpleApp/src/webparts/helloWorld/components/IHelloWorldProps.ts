import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as microsoftTeams from '@microsoft/teams-js';

export interface IHelloWorldProps {
  description: string;
  teamsContext: microsoftTeams.Context;
  webpartContext: WebPartContext;
}
