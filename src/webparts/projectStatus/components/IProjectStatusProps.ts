import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectStatusProps {
  hasTeamsContext: boolean;
  context: WebPartContext;
  backgroundImageUrl: string;
}
