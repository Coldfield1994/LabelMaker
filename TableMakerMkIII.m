function [Output] = TableMakerMkIII(Selection)
%This function creates a single label table. It doesn't fill in the table.

Table = Selection.Tables.Add(Selection.Range,7,7);
%Table.Borders.InsideLineStyle = 1;
%Table.Borders.OutsideLineStyle = 1;

Table.Rows.AllowBreakAcrossPages = 0;
Table.LeftPadding = 0;
Table.RightPadding = 0;

Table.Cell(1,1).Width = (1.3.*72);
Table.Cell(2,1).Width = (1.3.*72);
Table.Cell(3,1).Width = (1.3.*72);
Table.Cell(4,1).Width = (1.3.*72);
Table.Cell(5,1).Width = (1.3.*72);
Table.Cell(6,1).Width = (1.3.*72);
Table.Cell(7,1).Width = (1.3.*72);

Table.Cell(1,2).Width = (1.3.*72);
Table.Cell(2,2).Width = (1.3.*72);
Table.Cell(3,2).Width = (1.3.*72);
Table.Cell(4,2).Width = (1.3.*72);
Table.Cell(5,2).Width = (1.3.*72);
Table.Cell(6,2).Width = (1.3.*72);
Table.Cell(7,2).Width = (1.3.*72);

Table.Cell(1,3).Width = (0.9.*72);
Table.Cell(2,3).Width = (0.9.*72);
Table.Cell(3,3).Width = (0.9.*72);
Table.Cell(4,3).Width = (0.9.*72);
Table.Cell(5,3).Width = (0.9.*72);
Table.Cell(6,3).Width = (0.9.*72);
Table.Cell(7,3).Width = (0.9.*72);

Table.Cell(1,4).Width = (0.65.*72);
Table.Cell(2,4).Width = (0.65.*72);
Table.Cell(3,4).Width = (0.65.*72);
Table.Cell(4,4).Width = (0.65.*72);
Table.Cell(5,4).Width = (0.65.*72);
Table.Cell(6,4).Width = (0.65.*72);
Table.Cell(7,4).Width = (0.65.*72);

Table.Cell(1,5).Width = (1.3.*72);
Table.Cell(2,5).Width = (1.3.*72);
Table.Cell(3,5).Width = (1.3.*72);
Table.Cell(4,5).Width = (1.3.*72);
Table.Cell(5,5).Width = (1.3.*72);
Table.Cell(6,5).Width = (1.3.*72);
Table.Cell(7,5).Width = (1.3.*72);

Table.Cell(1,6).Width = (1.3.*72);
Table.Cell(2,6).Width = (1.3.*72);
Table.Cell(3,6).Width = (1.3.*72);
Table.Cell(4,6).Width = (1.3.*72);
Table.Cell(5,6).Width = (1.3.*72);
Table.Cell(6,6).Width = (1.3.*72);
Table.Cell(7,6).Width = (1.3.*72);

Table.Cell(1,7).Width = (0.9.*72);
Table.Cell(2,7).Width = (0.9.*72);
Table.Cell(3,7).Width = (0.9.*72);
Table.Cell(4,7).Width = (0.9.*72);
Table.Cell(5,7).Width = (0.9.*72);
Table.Cell(6,7).Width = (0.9.*72);
Table.Cell(7,7).Width = (0.9.*72);
            
Table.Rows.Height = (2.*72./7)-1;

        
Table.Rows.Item(1).Height = (2.*72./14);
Table.Rows.Item(2).Height = (2.*72./14);
Table.Rows.Item(3).Height = (2.*72./14);
Table.Rows.Item(4).Height = (2.*72./14);
Table.Rows.Item(5).Height = (2.*72./14);
Table.Rows.Item(6).Height = (2.*72.*8./14)-1;
Table.Rows.Item(7).Height = (2.*72./14);
            
%Table.Cell(5,1).Merge(Table.Cell(5,3));
Table.Cell(6,1).Merge(Table.Cell(6,3));
Table.Cell(7,1).Merge(Table.Cell(7,3));
%Table.Cell(1,3).Merge(Table.Cell(4,3));
Table.Cell(1,3).Merge(Table.Cell(5,3));

%Table.Cell(5,3).Merge(Table.Cell(5,5));
Table.Cell(6,3).Merge(Table.Cell(6,5));
Table.Cell(7,3).Merge(Table.Cell(7,5));
%Table.Cell(1,7).Merge(Table.Cell(4,7));
Table.Cell(1,7).Merge(Table.Cell(5,7));

Output = Table;

end