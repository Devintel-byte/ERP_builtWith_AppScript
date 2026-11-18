function requisitionClientDto_(requisition) {
  const clientRecord = Object.assign({}, requisition);
  delete clientRecord.token;
  return clientRecord;
}
