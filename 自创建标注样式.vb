'�Դ�����ע��ʽ1
Public Function AddDimStyle1()
    Dim dimStyle As AcadDimStyle
    Set dimStyle = ThisDrawing.DimStyles.Add("dimStyle1")
    ThisDrawing.ActiveDimStyle = dimStyle '����ñ�ע��ʽ
   
   With ThisDrawing
       '��һ�鶨��ȫ�ֺ����Ա�������
         .SetVariable "DimScale", 1     '����ȫ�ֱ�������
         .SetVariable "DimLFac", 1   '���Ա�������. '1'=1:1, '2'=2:1,'.5'=1:2��
        '������͵ı�ע����
        .SetVariable "DimADec", 0      '���ƽǶȱ�ע����ʾ��ȷλ��
        .SetVariable "DimAssoc", 2     '���Ʊ�ע����Ĺ�����
                                       'ʵ���ϸ�ϵͳ������ͼ�ο���
        .SetVariable "DimASz", 1.5        '���Ƴߴ��ߡ����߼�ͷ�Ĵ�С�������ƹ��ߵĴ�С
        .SetVariable "DimAtFit", 3    '���ߴ���ߵĿռ䲻����ͬʱ���±�ע���ֺͼ�ͷʱ,ȷ�������ߵ����з�ʽ
                                        '0 �����ֺͼ�ͷ�������ڳߴ����֮��
                                        '1  ���ƶ���ͷ��Ȼ���ƶ�����
                                        '2  ���ƶ����֣�Ȼ���ƶ���ͷ
                                        '3  �ƶ����ֺͼ�ͷ�нϺ��ʵ�һ��
        .SetVariable "DimAUnit", 0     '���ýǶȱ�ע�ĵ�λ��ʽ
                                       '0 ʮ���ƶ���
        .SetVariable "DimAZin", 0      '�ԽǶȱ�ע�����㴦��
                                       '0 ��ʾ����ǰ����ͺ�����
        .SetVariable "DimBlk", ""      '���óߴ��߻�����ĩ����ʾ�ļ�ͷ��
                                       '"" ʵ�ıպ�
        .SetVariable "DimBlk1", ""     '�� DIMSAH ϵͳ������ʱ�����óߴ��ߵ�һ���˵�ļ�ͷ
        .SetVariable "DimBlk2", ""     '�� DIMSAH ϵͳ������ʱ�����óߴ��ߵڶ����˵�ļ�ͷ
        .SetVariable "DimClrD", 256     'Ϊ�ߴ��ߡ���ͷ�ͱ�ע����ָ����ɫ
        .SetVariable "DimClrE", 256    'Ϊ�ߴ����ָ����ɫ������ɫ������������Ч����ɫ���
        .SetVariable "DimClrT", 256     'Ϊ��ע����ָ����ɫ
         .SetVariable "DimDec", 0       '���ñ�ע����λ��ʾ��С��λλ��
        .SetVariable "DimExe", 1        'ָ���ߴ���߳����ߴ��ߵľ���
        .SetVariable "DimExO", 6       'ָ���ߴ����ƫ��ԭ��ľ���
        .SetVariable "DimFrac", 0      '�� DIMLUNIT ϵͳ��������Ϊ 4���������� 5��������ʱ���÷�����ʽ
        .SetVariable "DimGap", 0.5     '���ߴ��߷ֳɶ���������֮����ñ�ע����ʱ�����ñ�ע������Χ�ľ���
        .SetVariable "DimJust", 0      '���Ʊ�ע���ֵ�ˮƽλ��
                                        '0  ���������ڳߴ���֮�ϣ����ڳߴ����֮�����ж���
                                        '1  ���ڵ�һ���ߴ���߷��ñ�ע����
                                        '2  ���ڵڶ����ߴ���߷��ñ�ע����
                                        '3  ����ע���ַ��ڵ�һ���ߴ�������ϣ�����֮����
                                        '4  ����ע���ַ��ڵڶ����ߴ�������ϣ�����֮����
        .SetVariable "DimLwd", acLnWtByLayer 'ָ���ߴ��ߵ��߿�
        .SetVariable "DimLwe", acLnWtByLayer 'ָ���ߴ���ߵ��߿�
        .SetVariable "DimPost", ""     'ָ����ע����ֵ������ǰ׺���׺���������߶�ָ����
        .SetVariable "DimRnd", 0       '�����б�ע�������뵽ָ��ֵ
        .SetVariable "DimSAh", 0       '���Ƴߴ��߼�ͷ�����ʾ
        .SetVariable "DimSD1", 0       '�����Ƿ��ֹ��ʾ��һ���ߴ���
        .SetVariable "DimSD2", 0       '�����Ƿ��ֹ��ʾ�ڶ����ߴ���
        .SetVariable "DimSE1", 0       '�����Ƿ��ֹ��ʾ��һ���ߴ����
        .SetVariable "DimSE2", 0       '�����Ƿ��ֹ��ʾ�ڶ����ߴ����
        .SetVariable "DimSOXD", 0      '�����Ƿ�����ߴ��߻��Ƶ��ߴ����֮��
        .SetVariable "DimTAD", 1       '����������Գߴ��ߵĴ�ֱλ��
                                       '0 ��ע�����ڳߴ����֮����з���
                                        '1  ���ǳߴ��߲���ˮƽ���õĻ��߳ߴ�����ڵ����ֱ�ǿ��Ϊˮƽ����
                                        '(DIMTIH = 1)������ͽ���ע���ַ����ڳߴ��ߵ��Ϸ�����ע������ײ�
                                        '���ߵ��ߴ��ߵľ���ֵ����ϵͳ����DIMGAP �ĵ�ǰֵ��
        .SetVariable "DimTIH", 0       '�������б�ע���ͣ������ע���⣩�ı�ע�����ڳߴ�����ڵ�λ��
                                        '0 ��� ��������ߴ��߶���
                                        '1 �� ������ˮƽ����
        .SetVariable "DimTIX", 1      '�ڳߴ����֮���������
                                        '0 ��� ������ע���͵Ĳ�ͬ����ͬ���������ԺͽǶȱ�ע��AutoCAD
                                        '�����ַ��õ��ߴ����֮�䣨������㹻�Ŀռ䣩�����ڲ����ڷ���Բ
                                        '��Բ���еİ뾶��ע��ֱ����ע��DIMTIX ��Ч������ǿ�ƽ����ַŵ�Բ��Բ��֮��
                                        '1 �� ����ע���ֻ����ڳߴ����֮�䣬��ʹ AutoCAD ͨ������Щ���ַ����ڳߴ����֮�⡣
        .SetVariable "DimTMOVE", 2      '���ñ�ע���ֵ��ƶ�����
                                        '0  �ߴ��ߺͱ�ע����һ���ƶ�
                                        '1  ���ƶ���ע����ʱ���һ������
                                        '2  �����ע���������ƶ��������������
        .SetVariable "DimTOFL", 0      '�����Ƿ񽫳ߴ��߻����ڳߴ����֮�䣨��ʹ���ַ����ڳߴ����֮�⣩
        .SetVariable "DimTOH", 0       '���Ʊ�ע�����ڳߴ�������λ��
        .SetVariable "DimTSz", 0      'ָ�����Ա�ע���뾶��ע�Լ�ֱ����ע�������ͷ��Сб�߳ߴ�
        .SetVariable "DimTVP", 0        '���Ƴߴ����Ϸ����·���ע���ֵĴ�ֱλ��
        .SetVariable "DimTxSty", "STANDARD"     'ָ����ע��������ʽ
        .SetVariable "DimTxt", 1.8         'ָ����ע���ֵĸ߶ȣ����ǵ�ǰ������ʽ���й̶��ĸ߶�
        .SetVariable "DimUPT", 0        '�����û���λ���ֵ�ѡ��
        .SetVariable "DimZIn", 0        '�����Ƿ������λֵ�����㴦��
'
        '���廻�㵥λ������
        .SetVariable "DimAlt", 0        '���Ʊ�ע�л��㵥λ����ʾ
        .SetVariable "DimAltD", 4       '���ƻ��㵥λ��С��λ��λ��
        .SetVariable "DimAltF", 25.4    '���ƻ��㵥λ����
        .SetVariable "DimAltRnd", 0     '���뻻���ע��λ
        .SetVariable "DimAltTD", 4      '���ñ�ע���㵥λ����ֵС��λ��λ��
        .SetVariable "DimAltTZ", 0      '�����Ƿ�Թ���ֵ�����㴦��
        .SetVariable "DimAltU", 2       'Ϊ���б�ע��ʽ�壨�Ƕȱ�ע���⣩���㵥λ���õ�λ��ʽ
        .SetVariable "DimAltZ", 0       '�����Ƿ�Ի��㵥λ��עֵ�����㴦��
        .SetVariable "DimAPost", ""     'Ϊ���б�ע���ͣ��Ƕȱ�ע���⣩�Ļ����ע����ֵָ������ǰ׺���׺�������߶�ָ����
   End With
    '��ע��ʽ�����Դ�ͼ��������ʽ�л��
   dimStyle.CopyFrom ThisDrawing
End Function