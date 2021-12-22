import { MSGraphClient } from '@microsoft/sp-http';
export interface IReactTestProps {
  description: string;
  graphClient: MSGraphClient;
}
