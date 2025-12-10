import React from 'react';
import RecordView from '../components/RecordView';
import SharePointService from '../services/sharepointService';

interface DetailScreenProps {
  sharePointService: SharePointService;
  listName: string;
  record: any;
  onBack: () => void;
  onRecordUpdated: () => void;
  employees?: any[];
}

const DetailScreen: React.FC<DetailScreenProps> = ({
  sharePointService,
  listName,
  record,
  onBack,
  onRecordUpdated,
  employees = [],
}) => {
  return (
    <RecordView
      sharePointService={sharePointService}
      listName={listName}
      record={record}
      onClose={onBack}
      onRecordUpdated={onRecordUpdated}
      onRecordDeleted={onRecordUpdated}
      employees={employees}
    />
  );
};

export default DetailScreen;
