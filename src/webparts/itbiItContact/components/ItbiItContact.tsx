import * as React from 'react';
import styles from './ItbiItContact.module.scss';
import { IItbiItContactProps } from './IItbiItContactProps';
import { IItbiItContactState } from './itbiITContactState';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import { Carousel, CarouselButtonsLocation, CarouselButtonsDisplay, CarouselIndicatorShape } from "@pnp/spfx-controls-react/lib/Carousel";
import { Icon } from 'office-ui-fabric-react/lib/Icon';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class ItbiItContact extends React.Component<IItbiItContactProps, IItbiItContactState> {

  constructor(props: IItbiItContactProps, state: IItbiItContactState) {
    super(props);
    sp.setup({
      spfxContext: this.props.context 
    });
    this.state = {
      carouselElements: []
    }; 
    this._getItems();   
  }

  private async _getItems() {
    const items: any[] = await sp.web.lists.getByTitle("Projects").items.get();
    let project: any[] = []; 
    let i: number;
    items.forEach(element => {
      i++;
      project.push(<div key={i} >
        <div className={ styles.wrapper }>                   
          <div className={ styles.icon }>
              <a href={"mailto:" + element.Link }>
                  <Icon iconName='Mail'/>
                  <span className={ styles.tooltip }>Отправить письмо</span>
              </a>
          </div>
          <div className={ styles.project }>  
              <div className={ styles.title }> 
                { element.Title }
              </div>
              <div className={ styles.mail }>    
                { element.Link }
              </div>
          </div>
        </div>
      </div>);       
    }); 
    this.setState({ carouselElements: project });    
  }

  public render(): React.ReactElement<IItbiItContactProps> {          
    return ( 
      <div className={ styles.itbiItContact }>
        <div className={ styles.container }>
          <div className={ styles.header }>
              Контактная информация поддержки IT
          </div>
          <div>
            <Carousel
              buttonsLocation={CarouselButtonsLocation.top}
              buttonsDisplay={CarouselButtonsDisplay.block}
              isInfinite={true}
              element={this.state.carouselElements}
              onMoveNextClicked={(index: number) => { }}
              onMovePrevClicked={(index: number) => { }} 
              pauseOnHover={true}
              containerStyles={styles.carousel}
              indicatorShape={CarouselIndicatorShape.circle}
              indicatorStyle={{ padding: 0, backgroundColor: "#333" }}   
              interval={60000}
            /> 
          </div>
        </div>
      </div>
    );
  }
}
