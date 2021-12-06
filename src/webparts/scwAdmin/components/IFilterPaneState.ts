import {IRequestItem} from './ScwAdmin';

export interface IFilterPaneState {
    isOpen: boolean;
    checkboxs: {
        checkbox1: boolean,
        checkbox2: boolean,
        checkbox3: boolean,
        checkbox4: boolean,
      };
      items: IRequestItem [];
} 