import * as React from 'react';

import 'jest';
import { createRenderer } from 'react-test-renderer/shallow';
import { shallow, mount } from 'enzyme';
import toJson from 'enzyme-to-json';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Sample, ISampleProps} from './Sample';

//jest.mock('@microsoft/sp-core-library/', () => 'BrowserDetection');
jest.mock('@microsoft/sp-core-library/', () => {
  return {
    'BrowserDetection': {},
    'DisplayMode': {}
  };
});
jest.mock('@pnp/spfx-controls-react/lib/WebPartTitle', () => {
  return {
    'WebPartTitle': 'WebPartTitle'
  };
});

test('should render Sample component correctly', () => {
  /*
   * using the OOTB 'react-test-renderer'
   */
  /*const renderer = createRenderer();
  renderer.render(<Sample />);
  console.log(renderer.getRenderOutput());
  expect(renderer.getRenderOutput()).toMatchSnapshot();*/

  /*
   * using enzyme
   */
  
  const wrapper = shallow(React.createElement(Sample, {}));
  expect(wrapper.find('span.message').text()).toBe('Hello world:');
  expect(wrapper.find('li').length).toBe(3);

  expect(toJson(wrapper)).toMatchSnapshot();
  expect(wrapper).toMatchSnapshot();
});