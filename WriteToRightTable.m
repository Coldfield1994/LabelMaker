function [] = WriteToRightTable(Selection, Table, Customer, Product, Counter, Counter2)
%This function writes a single lable data set to a single table.

% Right Table

%% Table Column 5, Rows 1-4

% Customer Name Length Check
CharacterArray = char(Customer.Name(Counter));
NameLength = length(CharacterArray);
if NameLength > 20
    LogicalArray = isspace(CharacterArray);
    HalfLength = ceil(NameLength./2);
    SelectorBool = false;
    Alternator = 1;
    FunctionCounter = 0;
    while SelectorBool == false
        if LogicalArray(HalfLength+FunctionCounter.*Alternator)
            SelectorBool = true;
            SpaceSelection = HalfLength+FunctionCounter.*Alternator;
            FirstLineName = extractBefore(CharacterArray, SpaceSelection);
            SecondLineName = extractAfter(CharacterArray, SpaceSelection);
        end
        if Alternator < 0
            FunctionCounter = FunctionCounter + 1;
            Alternator = -Alternator;
        else
            Alternator = -Alternator;
        end
    end

    Table.Cell(1,5).Range.Select;
    Table.Cell(1,5).Range.Font.Name = 'Arial';
    Table.Cell(1,5).Range.Font.Size = 6;
    Selection.TypeText(char(FirstLineName));
    
    Table.Cell(2,5).Range.Select;
    Table.Cell(2,5).Range.Font.Name = 'Arial';
    Table.Cell(2,5).Range.Font.Size = 6;
    Selection.TypeText(char(SecondLineName));

    Table.Cell(3,5).Range.Select;
    Table.Cell(3,5).Range.Font.Name = 'Arial';
    Table.Cell(3,5).Range.Font.Size = 6;
    Selection.TypeText(char(Customer.Address1(Counter)));

    Table.Cell(4,5).Range.Select;
    Table.Cell(4,5).Range.Font.Name = 'Arial';
    Table.Cell(4,5).Range.Font.Size = 6;
    Selection.TypeText(char(Customer.Address2(Counter)));

    Table.Cell(5,5).Range.Select;
    Table.Cell(5,5).Range.Font.Name = 'Arial';
    Table.Cell(5,5).Range.Font.Size = 6;
    Selection.TypeText(char(Customer.PhoneNumber(Counter)));
    
else
    FirstLineName = Customer.Name(Counter);
    
    Table.Cell(1,5).Range.Select;
    Table.Cell(1,5).Range.Font.Name = 'Arial';
    Table.Cell(1,5).Range.Font.Size = 6;
    Selection.TypeText(char(FirstLineName));

    Table.Cell(2,5).Range.Select;
    Table.Cell(2,5).Range.Font.Name = 'Arial';
    Table.Cell(2,5).Range.Font.Size = 6;
    Selection.TypeText(char(Customer.Address1(Counter)));

    Table.Cell(3,5).Range.Select;
    Table.Cell(3,5).Range.Font.Name = 'Arial';
    Table.Cell(3,5).Range.Font.Size = 6;
    Selection.TypeText(char(Customer.Address2(Counter)));

    Table.Cell(4,5).Range.Select;
    Table.Cell(4,5).Range.Font.Name = 'Arial';
    Table.Cell(4,5).Range.Font.Size = 6;
    Selection.TypeText(char(Customer.PhoneNumber(Counter)));
end

%% Table Column 6, Rows 1-4
Table.Cell(1,6).Range.Select;
Table.Cell(1,6).Range.Font.Bold = 1;
Table.Cell(1,6).Range.Font.Name = 'Arial';
Table.Cell(1,6).Range.Font.Size = 8;
Selection.TypeText(char(Customer.ProductNames(Counter,Counter2)));
Selection.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';

Table.Cell(2,6).Range.Select;
Table.Cell(2,6).Range.Font.Bold = 1;
Table.Cell(2,6).Range.Font.Name = 'Arial';
Table.Cell(2,6).Range.Font.Size = 8;
Selection.TypeText('Lot No. ');
Selection.TypeText(char(Product.LotNo(Counter2)));
Selection.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';

Table.Cell(3,6).Range.Select;
Table.Cell(3,6).Range.Font.Bold = 1;
Table.Cell(3,6).Range.Font.Name = 'Arial';
Table.Cell(3,6).Range.Font.Size = 8;
Selection.TypeText('Sp. Gr. = ')
Selection.TypeText(num2str(Product.SpecificGravity(Counter2)));
Selection.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';

Table.Cell(4,6).Range.Select;
Table.Cell(4,6).Range.Font.Bold = 1;
Table.Cell(4,6).Range.Font.Name = 'Arial';
Table.Cell(4,6).Range.Font.Size = 8;
Selection.TypeText('Exp. Date ');
Selection.TypeText(char(Product.ExpirationDate(Counter2)));
Selection.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';

%-------------------------------------------------------------------------

%% Table Column 3, Rows 1-4

Table.Cell(1,7).Range.Select;

invoke(Selection, 'TypeParagraph');
invoke(Selection, 'TypeParagraph');
invoke(Selection, 'TypeParagraph');
invoke(Selection, 'TypeParagraph');
invoke(Selection, 'TypeParagraph');
invoke(Selection, 'TypeParagraph');
invoke(Selection, 'TypeParagraph');

if (Product.Pictogram.ExclaimationMark(Counter2) && Product.Pictogram.Corrosion(Counter2))
    Pic = Selection.InLineShapes.AddPicture(fullfile(pwd, 'ExMark&CorrosivePictogram.jpg'));
elseif (Product.Pictogram.Corrosion(Counter2) && ~(Product.Pictogram.ExclaimationMark(Counter2)))
    Pic = Selection.InLineShapes.AddPicture(fullfile(pwd, 'CorrosivePictogram2.jpg'));
end
%--------------------------------------------------------------------------

%% Tabel Row 5
Table.Cell(5,6).Range.Select;
%Table.Cell(5,6).Range.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
Table.Cell(5,6).Range.Font.Bold = 1;
Table.Cell(5,6).Range.Font.Name = 'Arial';
Table.Cell(5,6).Range.Font.Size = 8;
Selection.TypeText('   ');
Selection.TypeText(char(Product.SignalWords(Counter2)));
Selection.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';

            
%% Tabel Row 6
Table.Cell(6,3).Range.Select;
Table.Cell(6,3).Range.Font.Name = 'Arial';
Table.Cell(6,3).Range.Font.Size = 6;
Selection.TypeText(char(Product.HazardStatement(Counter2)));
            
%% Tabel Row 7
Table.Cell(7,3).Range.Select;
Table.Cell(7,3).Range.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
Table.Cell(7,3).Range.Font.Name = 'Arial';
Table.Cell(7,3).Range.Font.Size = 6;
Selection.TypeText(char(Product.BeforeUseStatement(Counter2)));


end