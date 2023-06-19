def pre_annotation(talk_id):
    import pandas as pd

    pathway = '/Users/Foppevb/Documents/School/Studietijd/22-23/TW-3/Semester-2/Scriptie/'
    id = str(talk_id)
    filename_mdb = '/' + 'mdb_' + talk_id + '.xlsx'
    filename_q = '/' + 'q_' + talk_id + '.xlsx'

    mdb = pd.read_excel(pathway + id + filename_mdb)
    q = pd.read_excel(pathway + id + filename_q)

    mdb_objects = mdb.to_dict("records")
    q_objects = q.to_dict("records")

    annotation_double_ids = [
        {'chunk_start': q_dict['chunk_start'],
         'chunk_end': q_dict['chunk_end'],
         'highlight_start': q_dict['highlight_start'],
         'highlight_end': q_dict['highlight_end'],
         'wh-type': q_dict['wh_type'],
         'answered': q_dict['answered'],
         'related': q_dict['relatedness'],
         'q_id': q_dict['id'],
         'arg1': mdb_dict['Arg1'],
         'highlight': q_dict['highlight'],
         'question': q_dict['content'],
         'comment': '',
         'question_type': '',
         'relation_type': '',
         'repetition': '',
         'expected_sense1_l1': '',
         'expected_sense1_l2': '',
         'expected_sense1_l3': '',
         'expected_sense2_l1': '',
         'expected_sense2_l2': '',
         'expected_sense2_l3': '',
         'expected_sense3_l1': '',
         'expected_sense3_l2': '',
         'expected_sense3_l3': '',
         'arg2': mdb_dict['Arg2'],
         'sense1_l1': mdb_dict['l1_sense'],
         'sense1_l2': mdb_dict['l2_sense'],
         'sense1_l3': mdb_dict['l3_sense'],
         'relation_type': mdb_dict['Type']
         }
        for mdb_dict in mdb_objects
        for q_dict in q_objects
        if mdb_dict['Arg2_start'] >= q_dict['chunk_end']
        if mdb_dict['Arg1'] in str(q_dict['highlight'])
        or str(q_dict['highlight']) in mdb_dict['Arg1']
    ]

    ids = []
    annotation = []
    for dic in annotation_double_ids:
        if dic['q_id'] not in ids:
            annotation.append(dic)
        ids.append(dic['q_id'])

    df_annotation = pd.DataFrame.from_dict(annotation)

    df_annotation.to_excel(pathway + id + '/annotation_' + id + '.xlsx')

def post_annotation(talk_id):
    import pandas as pd

    pathway = '/Users/Foppevb/Documents/School/Studietijd/22-23/TW-3/Semester-2/Scriptie/'

    if talk_id == "total":
        annotated_1927 = pd.read_excel(pathway + '1927/annotation_1927.xlsx')
        annotated_1971 = pd.read_excel(pathway + '1971/annotation_1971.xlsx')
        annotated_1976 = pd.read_excel(pathway + '1976/annotation_1976.xlsx')

        annotated_1927_objects = annotated_1927.to_dict("records")
        annotated_1971_objects = annotated_1971.to_dict("records")
        annotated_1976_objects = annotated_1976.to_dict("records")

        alignment_1927 = [
            {'q_id': annotated_dict['q_id'],
             'arg1': annotated_dict['arg1'],
             'arg2': annotated_dict['arg2'],
             'l1_sense': annotated_dict['sense1_l1'],
             'l2_sense': annotated_dict['sense1_l2'],
             'l3_sense': annotated_dict['sense1_l3'],
             'expected_sense1_l1': annotated_dict['expected_sense1_l1'],
             'expected_sense1_l2': annotated_dict['expected_sense1_l2'],
             'expected_sense1_l3': annotated_dict['expected_sense1_l3'],
             'expected_sense2_l1': annotated_dict['expected_sense2_l1'],
             'expected_sense2_l2': annotated_dict['expected_sense2_l2'],
             'expected_sense2_l3': annotated_dict['expected_sense2_l3'],
             'expected_sense3_l1': annotated_dict['expected_sense3_l1'],
             'expected_sense3_l2': annotated_dict['expected_sense3_l2'],
             'expected_sense3_l3': annotated_dict['expected_sense3_l3'],
             'answered': annotated_dict['answered'],
             'related': annotated_dict['related'],
             'relation_type': annotated_dict['relation_type'],
             'relation_type_expected': annotated_dict['relation_type_expected'],
             'repetition': annotated_dict['repetition'],
             'question_type': annotated_dict['question_type']
             }
            for annotated_dict in annotated_1927_objects
        ]

        alignment_1971 = [
            {'q_id': annotated_dict['q_id'],
             'arg1': annotated_dict['arg1'],
             'arg2': annotated_dict['arg2'],
             'l1_sense': annotated_dict['sense1_l1'],
             'l2_sense': annotated_dict['sense1_l2'],
             'l3_sense': annotated_dict['sense1_l3'],
             'expected_sense1_l1': annotated_dict['expected_sense1_l1'],
             'expected_sense1_l2': annotated_dict['expected_sense1_l2'],
             'expected_sense1_l3': annotated_dict['expected_sense1_l3'],
             'expected_sense2_l1': annotated_dict['expected_sense2_l1'],
             'expected_sense2_l2': annotated_dict['expected_sense2_l2'],
             'expected_sense2_l3': annotated_dict['expected_sense2_l3'],
             'expected_sense3_l1': annotated_dict['expected_sense3_l1'],
             'expected_sense3_l2': annotated_dict['expected_sense3_l2'],
             'expected_sense3_l3': annotated_dict['expected_sense3_l3'],
             'answered': annotated_dict['answered'],
             'related': annotated_dict['related'],
             'relation_type': annotated_dict['relation_type'],
             'relation_type_expected': annotated_dict['relation_type_expected'],
             'repetition': annotated_dict['repetition'],
             'question_type': annotated_dict['question_type']
             }
            for annotated_dict in annotated_1971_objects
        ]

        alignment_1976 = [
            {'q_id': annotated_dict['q_id'],
             'arg1': annotated_dict['arg1'],
             'arg2': annotated_dict['arg2'],
             'l1_sense': annotated_dict['sense1_l1'],
             'l2_sense': annotated_dict['sense1_l2'],
             'l3_sense': annotated_dict['sense1_l3'],
             'expected_sense1_l1': annotated_dict['expected_sense1_l1'],
             'expected_sense1_l2': annotated_dict['expected_sense1_l2'],
             'expected_sense1_l3': annotated_dict['expected_sense1_l3'],
             'expected_sense2_l1': annotated_dict['expected_sense2_l1'],
             'expected_sense2_l2': annotated_dict['expected_sense2_l2'],
             'expected_sense2_l3': annotated_dict['expected_sense2_l3'],
             'expected_sense3_l1': annotated_dict['expected_sense3_l1'],
             'expected_sense3_l2': annotated_dict['expected_sense3_l2'],
             'expected_sense3_l3': annotated_dict['expected_sense3_l3'],
             'answered': annotated_dict['answered'],
             'related': annotated_dict['related'],
             'relation_type': annotated_dict['relation_type'],
             'relation_type_expected': annotated_dict['relation_type_expected'],
             'repetition': annotated_dict['repetition'],
             'question_type': annotated_dict['question_type']
             }
            for annotated_dict in annotated_1976_objects
        ]

        alignment = alignment_1927 + alignment_1971 + alignment_1976

        for dic in alignment:
            if dic['l1_sense'] == dic['expected_sense1_l1'] \
                    or dic['l1_sense'] == dic['expected_sense2_l1'] \
                    or dic['l1_sense'] == dic['expected_sense3_l1']:
                dic['alignment_l1'] = 1
            else:
                dic['alignment_l1'] = 0

            if dic['l2_sense'] == dic['expected_sense1_l2'] \
                    or dic['l2_sense'] == dic['expected_sense2_l2'] \
                    or dic['l2_sense'] == dic['expected_sense3_l2']:
                dic['alignment_l2'] = 1
            else:
                dic['alignment_l2'] = 0

            if dic['l3_sense'] == dic['expected_sense1_l3'] \
                    or dic['l3_sense'] == dic['expected_sense2_l3'] \
                    or dic['l3_sense'] == dic['expected_sense3_l3']:
                dic['alignment_l3'] = 1
            else:
                dic['alignment_l3'] = 0

            dic['alignment'] = dic['alignment_l1'] + dic['alignment_l2'] + dic['alignment_l3']

        for dic in alignment:
            dic['actual_sense'] = str(dic['l1_sense']) + '.' + str(dic['l2_sense']) + '.' + str(dic['l3_sense'])

        for dic in alignment:
            dic['sense_l1_l2'] = str(dic['l1_sense']) + '.' + str(dic['l2_sense'])

        for dic in alignment:
            dic['annotated_sense1'] = str(dic['expected_sense1_l1']) + '.' + str(dic['expected_sense1_l2']) + '.' + str(dic['expected_sense1_l3'])
            dic['annotated_sense2'] = str(dic['expected_sense2_l1']) + '.' + str(dic['expected_sense2_l2']) + '.' + str(dic['expected_sense2_l3'])
            dic['annotated_sense3'] = str(dic['expected_sense3_l1']) + '.' + str(dic['expected_sense3_l2']) + '.' + str(dic['expected_sense3_l3'])

        for dic in alignment:
            if dic['actual_sense'] == "Temporal.Synchronous.Synchronous":
                dic['actual_sense_abb'] = "Syn"
            elif dic['actual_sense'] == "Temporal.Asynchronous.Precedence":
                dic['actual_sense_abb'] = "Pre"
            elif dic['actual_sense'] == "Temporal.Asynchronous.Succession":
                dic['actual_sense_abb'] = "Suc"
            elif dic['actual_sense'] == "Contingency.Cause.Reason":
                dic['actual_sense_abb'] = "Rsn"
            elif dic['actual_sense'] == "Contingency.Cause.Result":
                dic['actual_sense_abb'] = "Rlt"
            elif dic['actual_sense'] == "Contingency.Cause.NegResult":
                dic['actual_sense_abb'] = "NRlt"
            elif dic['actual_sense'] == "Contingency.Cause+Belief.Reason+Belief":
                dic['actual_sense_abb'] = "RsnB"
            elif dic['actual_sense'] == "Contingency.Cause+Belief.Result+Belief":
                dic['actual_sense_abb'] = "RltB"
            elif dic['actual_sense'] == "Contingency.Cause+SpeechAct.Reason+SpeechAct":
                dic['actual_sense_abb'] = "RsnS"
            elif dic['actual_sense'] == "Contingency.Cause+SpeechAct.Result+SpeechAct":
                dic['actual_sense_abb'] = "RltS"
            elif dic['actual_sense'] == "Contingency.Condition.Arg1-as-cond":
                dic['actual_sense_abb'] = "1con"
            elif dic['actual_sense'] == "Contingency.Condition.Arg2-as-cond":
                dic['actual_sense_abb'] = "2con"
            elif dic['actual_sense'] == "Contingency.Condition+SpeechAct.Condition+SpeechAct":
                dic['actual_sense_abb'] = "CS"
            elif dic['actual_sense'] == "Contingency.Negative-Condition.Arg1-as-negCond":
                dic['actual_sense_abb'] = "1Ncon"
            elif dic['actual_sense'] == "Contingency.Negative-Condition.Arg2-as-negCond":
                dic['actual_sense_abb'] = "2Ncon"
            elif dic['actual_sense'] == "Contingency.Negative-Condition+SpeechAct.Negative-Condition+SpeechAct":
                dic['actual_sense_abb'] = "NgS"
            elif dic['actual_sense'] == "Contingency.Purpose.Arg1-as-goal":
                dic['actual_sense_abb'] = "1gl"
            elif dic['actual_sense'] == "Contingency.Purpose.Arg2-as-goal":
                dic['actual_sense_abb'] = "2gl"
            elif dic['actual_sense'] == "Comparison.Concession.Arg1-as-denier":
                dic['actual_sense_abb'] = "1den"
            elif dic['actual_sense'] == "Comparison.Concession.Arg2-as-denier":
                dic['actual_sense_abb'] = "2den"
            elif dic['actual_sense'] == "Comparison.Concession+SpeechAct.Arg2-as-denier+SpeechAct":
                dic['actual_sense_abb'] = "2denS"
            elif dic['actual_sense'] == "Comparison.Contrast.Contrast":
                dic['actual_sense_abb'] = "Ctr"
            elif dic['actual_sense'] == "Comparison.Similarity.Similarity":
                dic['actual_sense_abb'] = "Sim"
            elif dic['actual_sense'] == "Expansion.Conjunction.Conjunction":
                dic['actual_sense_abb'] = "Cnj"
            elif dic['actual_sense'] == "Expansion.Disjunction.Disjunction":
                dic['actual_sense_abb'] = "Dsj"
            elif dic['actual_sense'] == "Expansion.Equivalence.Equivalence":
                dic['actual_sense_abb'] = "Eqv"
            elif dic['actual_sense'] == "Expansion.Exception.Arg1-as-excpt":
                dic['actual_sense_abb'] = "1exc"
            elif dic['actual_sense'] == "Expansion.Exception.Arg2-as-excpt":
                dic['actual_sense_abb'] = "2exc"
            elif dic['actual_sense'] == "Expansion.Instantiation.Arg1-as-instance":
                dic['actual_sense_abb'] = "1ins"
            elif dic['actual_sense'] == "Expansion.Instantiation.Arg2-as-instance":
                dic['actual_sense_abb'] = "2ins"
            elif dic['actual_sense'] == "Expansion.Instantiation.Instantiation":
                dic['actual_sense_abb'] = "Ins"
            elif dic['actual_sense'] == "Expansion.Level-of-detail.Arg1-as-detail":
                dic['actual_sense_abb'] = "1det"
            elif dic['actual_sense'] == "Expansion.Level-of-detail.Arg2-as-detail":
                dic['actual_sense_abb'] = "2det"
            elif dic['actual_sense'] == "Expansion.Manner.Arg1-as-manner":
                dic['actual_sense_abb'] = "1man"
            elif dic['actual_sense'] == "Expansion.Manner.Arg2-as-manner":
                dic['actual_sense_abb'] = "2man"
            elif dic['actual_sense'] == "Expansion.Substitution.Arg1-as-subst":
                dic['actual_sense_abb'] = "1sub"
            elif dic['actual_sense'] == "Expansion.Substitution.Arg2-as-subst":
                dic['actual_sense_abb'] = "2sub"
            elif dic['actual_sense'] == "nan.nan.nan":
                dic['actual_sense_abb'] = "nan"
            else:
                dic['actual_sense_abb'] = dic['actual_sense']

        for dic in alignment:
            if dic['sense_l1_l2'] == "Temporal.Synchronous":
                dic['l2_abb'] = "Synch"
            elif dic['sense_l1_l2'] == "Temporal.Asynchronous":
                dic['l2_abb'] = "Asynch"
            elif dic['sense_l1_l2'] == "Contingency.Cause":
                dic['l2_abb'] = "Cause"
            elif dic['sense_l1_l2'] == "Contingency.Cause+Belief":
                dic['l2_abb'] = "Cause+B"
            elif dic['sense_l1_l2'] == "Contingency.Cause+SpeechAct":
                dic['l2_abb'] = "Cause+S"
            elif dic['sense_l1_l2'] == "Contingency.Condition":
                dic['l2_abb'] = "Cond"
            elif dic['sense_l1_l2'] == "Contingency.Condition+SpeechAct":
                dic['l2_abb'] = "Cond+S"
            elif dic['sense_l1_l2'] == "Contingency.Negative-condition":
                dic['l2_abb'] = "NegCond"
            elif dic['sense_l1_l2'] == "Contingency.Negative-condition+SpeechAct":
                dic['l2_abb'] = "NegCond+S"
            elif dic['sense_l1_l2'] == "Contingency.Purpose":
                dic['l2_abb'] = "Purpose"
            elif dic['sense_l1_l2'] == "Comparison.Concession":
                dic['l2_abb'] = "Conces"
            elif dic['sense_l1_l2'] == "Comparison.Concession+SpeechAct":
                dic['l2_abb'] = "Conces+S"
            elif dic['sense_l1_l2'] == "Comparison.Contrast":
                dic['l2_abb'] = "Contr"
            elif dic['sense_l1_l2'] == "Comparison.Similarity":
                dic['l2_abb'] = "Simil"
            elif dic['sense_l1_l2'] == "Expansion.Conjunction":
                dic['l2_abb'] = "Conj"
            elif dic['sense_l1_l2'] == "Expansion.Disjunction":
                dic['l2_abb'] = "Disj"
            elif dic['sense_l1_l2'] == "Expansion.Equivalence":
                dic['l2_abb'] = "Equiv"
            elif dic['sense_l1_l2'] == "Expansion.Exception":
                dic['l2_abb'] = "Excpt"
            elif dic['sense_l1_l2'] == "Expansion.Instantiation":
                dic['l2_abb'] = "Instant"
            elif dic['sense_l1_l2'] == "Expansion.Level-of-detail":
                dic['l2_abb'] = "LvlDet"
            elif dic['sense_l1_l2'] == "Expansion.Manner":
                dic['l2_abb'] = "Manner"
            elif dic['sense_l1_l2'] == "Expansion.Substitution":
                dic['l2_abb'] = "Subst"
            else:
                dic['l2_abb'] = dic['sense_l1_l2']

        R = [
            {'q_id': aligned_dict['q_id'],
             'answered': aligned_dict['answered'],
             'related': aligned_dict['related'],
             'alignment_l1': aligned_dict['alignment_l1'],
             'alignment_l2': aligned_dict['alignment_l2'],
             'alignment_l3': aligned_dict['alignment_l3'],
             'alignment': aligned_dict['alignment'],
             'actual_sense': aligned_dict['actual_sense'],
             'sense_l1_l2': aligned_dict['sense_l1_l2'],
             'abb': aligned_dict['actual_sense_abb'],
             'l2_abb': aligned_dict['l2_abb'],
             'sense_l1': aligned_dict['l1_sense'],
             'sense_l2': aligned_dict['l2_sense'],
             'sense_l3': aligned_dict['l3_sense'],
             'annotated_sense1': aligned_dict['annotated_sense1'],
             'annotated_sense2': aligned_dict['annotated_sense2'],
             'annotated_sense3': aligned_dict['annotated_sense3'],
             'relation_type': aligned_dict['relation_type'],
             'relation_type_expected': aligned_dict['relation_type_expected'],
             'repetition': aligned_dict['repetition'],
             'question_type': aligned_dict['question_type']
             }
            for aligned_dict in alignment
        ]

        df_R = pd.DataFrame.from_dict(R)

        df_R.to_excel(pathway + 'Total/R_Total.xlsx')

    else:
        id = str(talk_id)

        annotated = pd.read_excel(pathway + id + '/annotation_' + id + '.xlsx')

        annotated_objects = annotated.to_dict("records")

        alignment = [
            {'q_id': annotated_dict['q_id'],
             'arg1': annotated_dict['arg1'],
             'arg2': annotated_dict['arg2'],
             'l1_sense': annotated_dict['sense1_l1'],
             'l2_sense': annotated_dict['sense1_l2'],
             'l3_sense': annotated_dict['sense1_l3'],
             'expected_sense1_l1': annotated_dict['expected_sense1_l1'],
             'expected_sense1_l2': annotated_dict['expected_sense1_l2'],
             'expected_sense1_l3': annotated_dict['expected_sense1_l3'],
             'expected_sense2_l1': annotated_dict['expected_sense2_l1'],
             'expected_sense2_l2': annotated_dict['expected_sense2_l2'],
             'expected_sense2_l3': annotated_dict['expected_sense2_l3'],
             'expected_sense3_l1': annotated_dict['expected_sense3_l1'],
             'expected_sense3_l2': annotated_dict['expected_sense3_l2'],
             'expected_sense3_l3': annotated_dict['expected_sense3_l3'],
             'answered': annotated_dict['answered'],
             'related': annotated_dict['related'],
             'relation_type': annotated_dict['relation_type'],
             'relation_type_expected': annotated_dict['relation_type_expected'],
             'repetition': annotated_dict['repetition'],
             'question_type': annotated_dict['question_type']
             }
            for annotated_dict in annotated_objects
        ]

    for dic in alignment:
        if dic['l1_sense'] == dic['expected_sense1_l1'] \
                or dic['l1_sense'] == dic['expected_sense2_l1'] \
                or dic['l1_sense'] == dic['expected_sense3_l1']:
            dic['alignment_l1'] = 1
        else:
            dic['alignment_l1'] = 0

        if dic['l2_sense'] == dic['expected_sense1_l1'] \
                or dic['l2_sense'] == dic['expected_sense2_l2'] \
                or dic['l2_sense'] == dic['expected_sense3_l2']:
            dic['alignment_l2'] = 1
        else:
            dic['alignment_l2'] = 0

        if dic['l3_sense'] == dic['expected_sense1_l3'] \
                or dic['l3_sense'] == dic['expected_sense2_l3'] \
                or dic['l3_sense'] == dic['expected_sense3_l3']:
            dic['alignment_l3'] = 1
        else:
            dic['alignment_l3'] = 0

        dic['alignment'] = dic['alignment_l1'] + dic['alignment_l2'] + dic['alignment_l3']

    for dic in alignment:
        dic['actual_sense'] = str(dic['l1_sense']) + '.' + str(dic['l2_sense']) + '.' + str(dic['l3_sense'])

    for dic in alignment:
        dic['annotated_sense1'] = str(dic['expected_sense1_l1']) + '.' + str(dic['expected_sense1_l2']) + '.' + str(dic['expected_sense1_l3'])
        dic['annotated_sense2'] = str(dic['expected_sense2_l1']) + '.' + str(dic['expected_sense2_l2']) + '.' + str(dic['expected_sense2_l3'])
        dic['annotated_sense3'] = str(dic['expected_sense3_l1']) + '.' + str(dic['expected_sense3_l2']) + '.' + str(dic['expected_sense3_l3'])

    R = [
        {'q_id': aligned_dict['q_id'],
         'answered': aligned_dict['answered'],
         'related': aligned_dict['related'],
         'alignment_l1': aligned_dict['alignment_l1'],
         'alignment_l2': aligned_dict['alignment_l2'],
         'alignment_l3': aligned_dict['alignment_l3'],
         'alignment': aligned_dict['alignment'],
         'actual_sense': aligned_dict['actual_sense'],
         'relation_type': aligned_dict['relation_type'],
         'relation_type_expected': aligned_dict['relation_type_expected'],
         'repetition': aligned_dict['repetition'],
         'question_type': aligned_dict['question_type']
         }
        for aligned_dict in alignment
    ]

    df_R = pd.DataFrame.from_dict(R)

    df_R.to_excel(pathway + id + '/R_' + id + '.xlsx')
