import { IOwner } from './IOwner';
import {IRequestItem} from './ScwAdmin';
import {IContextualMenuProps} from '@fluentui/react';

export interface IScwAdminState {
    showDetailScreen: boolean;
    selectedTitle: string;
    selectedDesEn: string;
    selectedDesFr: string;
    selectedReason: string;
    selectedKey: string;
    selectedUrl: string;
    itemId: number;
    owners: string[];
    itemStatus: string;
    items: IRequestItem [];
    contextualMenuProps?: IContextualMenuProps;
    isPaneOpen: boolean;
    selectedRequesterName: string;
    selectedRequesterEmail: string;
}