import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IOwner } from './IOwner';

export interface IItemDetailProps {
    returnToMainPage: () => void;
    context: WebPartContext;
    spaceName: string;
    descriptionEn: string;
    descriptionFR: string;
    reason: string;
    itemId: number;
    owners: string[];
    itemStatus:string;
    url:string;
    requesterName: string;
    requesterEmail: string;
}