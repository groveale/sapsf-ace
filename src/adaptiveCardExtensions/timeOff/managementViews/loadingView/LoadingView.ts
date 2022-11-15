import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import { ITimeOffAdaptiveCardExtensionProps, ITimeOffAdaptiveCardExtensionState } from '../../TimeOffAdaptiveCardExtension';

export interface ILoadingViewData {
    title: string;
}

export class LoadingView extends BaseAdaptiveCardView<
ITimeOffAdaptiveCardExtensionProps,
ITimeOffAdaptiveCardExtensionState, ILoadingViewData> {
    public get data(): ILoadingViewData {
        return {
            title: `Request is in progress....`,
        };
    }
    
    public get template(): ISPFxAdaptiveCard {
        return require('./template/LoadingViewTemplate.json');
    }
}