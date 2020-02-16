import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { Nav, INavLinkGroup, INavLink } from 'office-ui-fabric-react';
//import { Placeholder } from "@pnp/spfx-controls-react";
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { SearchBox } from 'office-ui-fabric-react';

export interface ISampleProps {
}

export class Sample extends React.Component<ISampleProps, {}> {
  public render(): React.ReactElement<ISampleProps> {

    return (
      <div>
        <span className="message">Hello world:</span>
        <ul>
          <li>one</li>
          <li>two</li>
          <li>three</li>
        </ul>
        <MessageBar messageBarType={MessageBarType.info}>No results found.</MessageBar>
        <SearchBox />
        <Placeholder iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part.'
          buttonLabel='Configure' />
        <WebPartTitle title='Hello' updateProperty={() => { }} displayMode={DisplayMode.Read} />
        <Nav groups={[{links:[]}]}
          collapsedStateText="Collapsed"
          expandedStateText="Expand"
          key={Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15)}
        />
      </div>
    );
  }
}