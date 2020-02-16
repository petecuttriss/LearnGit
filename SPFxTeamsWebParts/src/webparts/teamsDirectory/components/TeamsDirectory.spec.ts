import * as React from 'react';
import 'jest';
import { shallow, mount, ReactWrapper, ShallowWrapper } from 'enzyme';
import TeamsDirectory from './TeamsDirectory';
import { MockListService } from '../../../services/MockListService';
import { IListService } from '../../../models/IListService';
import { DisplayMode } from "@microsoft/sp-core-library";
import { ITeamsDirectoryProps } from './ITeamsDirectoryProps';
import { ITeamsDirectoryState } from './ITeamsDirectoryState';

jest.mock('@microsoft/sp-core-library/', () => { return { 'BrowserDetection': {}, 'DisplayMode': {} }; });
jest.mock('@pnp/spfx-controls-react/lib/WebPartTitle', () => { return { 'WebPartTitle': 'WebPartTitle' }; });

describe('TeamsDirectory', () => {
  let wrapper: ShallowWrapper<ITeamsDirectoryProps, ITeamsDirectoryState>;
  let spy: jest.SpyInstance;
  
  beforeEach(() => {
    spy = jest.spyOn(TeamsDirectory.prototype, 'componentDidMount');
    wrapper = shallow(React.createElement(TeamsDirectory, {
      title: "Teams Directory",
      siteUrl: "/",
      listId: "1",
      hideTitle: false,
      hideSearchBox: false,
      openLinksInANewTab: true,
      displayMode: DisplayMode.Read,
      listService: new MockListService(),
      updateProperty: () => { },
      configureStartCallback: () => { }
    }));
  });

  test('should exist', () => {
    //const listServiceInstance: IListService = new MockListService();
    //const spy = jest.spyOn(TeamsDirectory.prototype, 'componentDidMount');
    /*const wrapper = shallow(React.createElement(TeamsDirectory, {
      title: "Teams Directory",
      siteUrl: "/",
      listId: "1",
      hideTitle: false,
      hideSearchBox: false,
      openLinksInANewTab: true,
      displayMode: DisplayMode.Read,
      listService: listServiceInstance,
      updateProperty: () => { },
      configureStartCallback: () => { }
    }));*/

    /*wrapper.setState({
      loading:false,
      error: null,
      isSearch: false,
      items: [
        {
          "odata.type": "SP.Data.SitesRegisterItem",
          "odata.id": "c805c25b-fef1-4273-97c8-11e71c7d86d8",
          "odata.etag": "\"4\"",
          "odata.editLink": "Web/Lists(guid'2b309a30-4a12-4167-a143-750601cffada')/Items(114)",
          "ContentType@odata.navigationLinkUrl": "Web/Lists(guid'2b309a30-4a12-4167-a143-750601cffada')/Items(114)/ContentType",
          "Title": "Auckland University Open Day",
          "SFDesc": null,
          "SIAStatus": "03 Active",
          "SFLink": {
            "Description": "Auckland University Open Day",
            "Url": "https://iworkplace.sharepoint.com/sites/AKLUNI"
          },
          "Function": "Corporate Management",
          "Activity": "Information Services",
          "Subactivity": null
        }]
    })*/
    expect(spy).toBeCalled();
    expect(wrapper).toMatchSnapshot();
    expect(wrapper.find('WebPartTitle').prop('title')).toBe("Teams Directory");
  });
});