import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react';

export interface IRegionFilterProps {
  regions: string[]; // distinct regions present in the employee list, e.g. ['South Africa', 'UK', 'US']
  selectedRegion: string | undefined; // undefined = "All"
  onChange: (region: string | undefined) => void;
}

const ALL_KEY = '__all__';

// Region filter (All / South Africa / UK / US) narrows the employee picker
// before showing options - confirmed design rule. Regions are read from
// the live employee data rather than hardcoded, so this adapts if the
// real list uses different region labels than expected.
const RegionFilter: React.FC<IRegionFilterProps> = ({ regions, selectedRegion, onChange }) => {
  return (
    <Pivot
      selectedKey={selectedRegion || ALL_KEY}
      onLinkClick={(item?: PivotItem) => {
        if (!item) return;
        onChange(item.props.itemKey === ALL_KEY ? undefined : item.props.itemKey);
      }}
    >
      <PivotItem headerText="All" itemKey={ALL_KEY} />
      {regions.map(region => (
        <PivotItem headerText={region} itemKey={region} key={region} />
      ))}
    </Pivot>
  );
};

export default RegionFilter;
