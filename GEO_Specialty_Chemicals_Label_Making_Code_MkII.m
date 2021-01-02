%GEO Specialty Chemicals Label-Making Code.
%Author: Christopher Oldfield
%Edition I
%Edited 11/17/2019

%This code can handle a maximum of 35 products at once.

%Helpful Website:
%https://docs.microsoft.com/en-us/visualstudio/vsto/word-object-model-overview?view=vs-2019

%% Section 1 - Initilizations
% This section initilizes variables, loads excel data, and turns on necessary applications.

% Initializers
Counter = [];
Counter2 = [];
Counter3 = [];
TableParse = [];
InToPoints = 72;

% Initializing MS Word
Word = actxserver('Word.Application');
Word.Visible = 1;
Document = Word.Documents.Add;
Selection = Word.Selection;
PageSetup = Selection.PageSetup;
ParagraphFormat = Selection.ParagraphFormat;
Selection.Font.Size = 6;
ParagraphFormat.LineSpacingRule = 'wdLineSpaceExactly';
ParagraphFormat.LineSpacing = 6;
ParagraphFormat.SpaceAfter = 0;
ParagraphFormat.WidowControl = -1;


% Declaring Excel Spreadsheet
[Excel.File,Excel.Path] = uigetfile({'*.xlsx'},'Excel File Selector');
if isequal(Excel.File,0)
    disp('File not selected.');
    return;
end



%% Section 2 - Data Manipulation
% This section takes loaded excel data and converts it to MATLAB data files. 

% Reading Excel Spreadsheet
[Data.Num,Data.Text,Data.Raw] = xlsread([Excel.Path,Excel.File]);
    % Cropping Extra Y-Axis Data
    Logical = cellfun(@isnan, Data.Raw(:,1),'UniformOutput',false);
    Logical = cell2mat(cellfun(@(x) x(1), Logical,'UniformOutput',false));
    for Counter = 1:length(Logical)
        if Logical(Counter)
            Excel.Y_Length = Counter-1;
            Customer.NumberSansGEO = Excel.Y_Length-15;
            Customer.NumberPlusGEO = Excel.Y_Length-14;
            Data.Raw = Data.Raw(1:Counter-1,:);
            break
        else
            Excel.Y_Length = Counter;
            Customer.NumberSansGEO = Excel.Y_Length-15;
            Customer.NumberPlusGEO = Excel.Y_Length-14;
        end
    end

    clear('Logical','Counter');

% Reorganizing Excel Data

% Products
Product.Name = Data.Raw(4,2:36);
Product.Description = Data.Raw(5,2:36);
Product.DistributionTime = Data.Raw(3,2:36);
Product.LotNo = Data.Raw(end-8,2:36);
Product.SpecificGravity = cell2mat(Data.Raw(end-7,2:36));
Product.PercentAl2O3 = Data.Raw(end-6,2:36);
    for Counter = 1:length(Product.PercentAl2O3)
        if isequal(Product.PercentAl2O3(Counter),{'NA'})
            Product.PercentAl2O3(Counter) = {0};
        end
    end
    Counter = [];
    Product.PercentAl2O3 = cell2mat(Product.PercentAl2O3);
Product.ExpirationDate = Data.Raw(end-5,2:36);
Product.SignalWords = Data.Raw(end-4,2:36);

    Logical = cellfun(@isnan, Product.SignalWords,'UniformOutput',false);
    Logical = cell2mat(cellfun(@(x) x(1), Logical,'UniformOutput',false));
    Product.SignalWords(Logical) = {' '};
    clear('Logical');
    
        
Product.HazardStatement = Data.Raw(end-3,2:36);
Product.BeforeUseStatement = Data.Raw(end-2,2:36);
Product.CustomerProductNames = Data.Raw(6:end-9,2:end-5);
Product.Pictogram.ExclaimationMark = cell2mat(Data.Raw(end-1,2:36));
Product.Pictogram.Corrosion = cell2mat(Data.Raw(end,2:36));

% Customers
Customer.Name = Data.Raw(6:end-9,37);
Customer.Address1 = Data.Raw(6:end-9,38);
Customer.Address2 = Data.Raw(6:end-9,39);
Customer.PhoneNumber = Data.Raw(6:end-9,40);
Customer.ProductNames = Data.Raw(6:end-9,2:end-5);
Customer.LabelNum = cell2mat(Data.Raw(6:end-9,41));


%% Section 3 - Label Printing
% This section takes MATLAB data and prints it to MS Word for label-making.

% Word Document Initial Setup
PageSetup.TopMargin = 0.6.*InToPoints;
PageSetup.BottomMargin = 0.3.*InToPoints;
PageSetup.LeftMargin = (0.16+0.05+0.1).*InToPoints;
PageSetup.RightMargin = (0.16+0.05).*InToPoints;
Selection.Paragraphs.LineUnitAfter = 0;

% Half-Year Bool
HalfYear = false;


% Label Creation

LabelParse = 1;
TableParse = 1;

for Counter = 1:Customer.NumberPlusGEO
    
    % No Labels Check
    if ~(isequal(Customer.LabelNum(Counter),0))
        
        for Counter2 = 1:Customer.LabelNum(Counter)

            for Counter3 = 1:35         

                % Label Existence Check
                if isequal(Product.CustomerProductNames(Counter,Counter3),{'DELETE'})
                    
                    % End of Set Check
                    if (Counter3 == 35) && (Counter2 == Customer.LabelNum(Counter)) && (Counter ~= Customer.NumberPlusGEO)
                        PageCheck = rem(LabelParse,10);
                        if PageCheck ~= 0
                            invoke(Selection, 'MoveRight');
                            invoke(Selection, 'MoveDown');
                            Selection.InsertNewPage;
                            Selection.Extend;
                            Selection = Word.Selection;
                            PageSetup = Selection.PageSetup;
                            ParagraphFormat = Selection.ParagraphFormat;
                            Selection.Font.Size = 6;
                            ParagraphFormat.LineSpacingRule = 'wdLineSpaceExactly';
                            ParagraphFormat.LineSpacing = 6;
                            ParagraphFormat.SpaceAfter = 0;
                            ParagraphFormat.WidowControl = -1;
                            Selection = Word.Selection;
                            LabelParse = 1;
                        end
                    end
                    
                    continue;
                end
                
                % Label Half-Year Check
                if HalfYear
                    if isequal(Product.DistributionTime(Counter3),{'12 months'})
                        
                        % End of Set Check
                        if (Counter3 == 35) && (Counter2 == Customer.LabelNum(Counter)) && (Counter ~= Customer.NumberPlusGEO)
                            PageCheck = rem(LabelParse,10);
                            if PageCheck ~= 0
                                invoke(Selection, 'MoveRight');
                                invoke(Selection, 'MoveDown');
                                Selection.InsertNewPage;
                                Selection.Extend;
                                Selection = Word.Selection;
                                PageSetup = Selection.PageSetup;
                                ParagraphFormat = Selection.ParagraphFormat;
                                Selection.Font.Size = 6;
                                ParagraphFormat.LineSpacingRule = 'wdLineSpaceExactly';
                                ParagraphFormat.LineSpacing = 6;
                                ParagraphFormat.SpaceAfter = 0;
                                ParagraphFormat.WidowControl = -1;
                                Selection = Word.Selection;
                                LabelParse = 1;
                            end
                        end
                        
                        continue;
                    end
                end

                % Left or Right Label Check
                if mod(LabelParse,2) %Left Label
                    % Table Creation
                    Table(TableParse) = TableMakerMkIII(Selection);
                    WriteToLeftTable(Selection, Table(TableParse), Customer, Product, Counter, Counter3);

                else %Right Label
                    WriteToRightTable(Selection, Table(TableParse), Customer, Product, Counter, Counter3);

                    % New Page Check
                    PageCheck = rem(LabelParse,10);
                    if PageCheck ~= 0
                        invoke(Selection, 'MoveDown');
                        ParagraphFormat.LineSpacing = 1;
                        invoke(Selection, 'TypeParagraph');
                        ParagraphFormat.LineSpacing = 6;

                    else
                        invoke(Selection, 'MoveRight');
                        invoke(Selection, 'MoveDown');
                        Selection.InsertNewPage;
                        Selection.Extend;
                        Selection.MoveUp;
                        Selection.Delete;
                        Selection = Word.Selection;
                        PageSetup = Selection.PageSetup;
                        ParagraphFormat = Selection.ParagraphFormat;
                        Selection.Font.Size = 6;
                        ParagraphFormat.LineSpacingRule = 'wdLineSpaceExactly';
                        ParagraphFormat.LineSpacing = 6;
                        ParagraphFormat.SpaceAfter = 0;
                        ParagraphFormat.WidowControl = -1;
                        Selection = Word.Selection;

                    end
                    TableParse = TableParse+1;

                end
                
                % End of Set Check
                if (Counter3 == 35) && (Counter2 == Customer.LabelNum(Counter)) && (Counter ~= Customer.NumberPlusGEO)
                    PageCheck = rem(LabelParse,10);
                    if PageCheck ~= 0
                        invoke(Selection, 'MoveRight');
                        invoke(Selection, 'MoveDown');
                        Selection.InsertNewPage;
                        Selection.Extend;
                        Selection = Word.Selection;
                        PageSetup = Selection.PageSetup;
                        ParagraphFormat = Selection.ParagraphFormat;
                        Selection.Font.Size = 6;
                        ParagraphFormat.LineSpacingRule = 'wdLineSpaceExactly';
                        ParagraphFormat.LineSpacing = 6;
                        ParagraphFormat.SpaceAfter = 0;
                        ParagraphFormat.WidowControl = -1;
                        Selection = Word.Selection;
                        LabelParse = 0;
                    end
                end
                LabelParse = LabelParse+1;

            end
            %Counter3 = 1;
        end
    end
    %Counter2 = 1;
    
end

%clear('Counter', 'Counter2', 'Counter3', 'LabelParse', 'TableParse', 'PageCheck');





%% Section 4 - Ending Application
% This section saves and closes all relavent data and applications.

%Document.SaveAs2([pwd,'/test.docx']);
%Word.Quit();






%% Section 5 Extra Data For Coding

%{
 WdUnits.
    wdCell = 12  
    wdCharacter = 1
    wdCharacterFormatting = 13
    wdColumn = 9
    wdItem = 16
    wdLine = 5
    wdParagraph = 4
    wdParagraphFormatting = 14
    wdRow = 10
    wdScreen = 7
    wdSection = 8
    wdSentence = 3
    wdStory = 6
    wdTable = 15
    wdWindow = 11
    wdWord = 2
 WdMovementType.
    wdExtend = 1
    wWdMove = 0
%}
