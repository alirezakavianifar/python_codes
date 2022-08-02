import pandas as pd
import os
import time
import glob

s_path = r'D:\SOFTWARE\db\downloads\based on words\s_source.xlsx'
sf_path = r'D:\SOFTWARE\db\downloads\based on words\s_source filtered by abstract.xlsx'

searchfor = ['self-adaptive', 'adaptation','three layer','unanticipated','Dynamically Discovering','planning at run time','automated planning',
             'adaptive planning','adaptive plan','adaptive feedback','adaptive monitor','adaptable monitor','dynamic environment','MAPE-K'
             'runtime evolution','evolutionary','search-based','genetic programming','optimizing feedback loop','context aware'
             ,'adapting analyzer','self-healing','self-optimizing','self-learning','self-protecting','self-improving','automated capabilities',
             'self-manag','enexpected change','changing behaviour','meta-manager','reusable self-adaptive','reuse','automatic optimization techniques',
             'meta-interface','higher order system','meta layer','new adaptation strategies','unforeseen change',
             'higher-order adaptation','at runtime','self-awareness','changes in behavior','change and uncertainty',
             'changes to the adaptive system','unexpected situation','reuses existing plan','reuse existing plan','reusing',
             'handling unanticipated changes','known unknowns','second-order','self-* planning','replanning for unexpected','Meta-self-awareness',
             'self-adjust','runtime adjustment','online planner','plans at runtime','autonomous manager','governing policies','survey',
             'adapt the behavior of self-adaptive','unanticipated adaptation','evolving self-adaptive','Managing Uncertainty','Systematic Mapping']

sources=['Software Engineering for Adaptive','Self-Adaptive and Self-Organizing',
         'Transactions on Autonomous','International Conference on Autonomic',
         'European Software Engineering Conference','Workshop on Self-healing',
         'International Conference on Software Engineering',
         'Transactions on Software Engineering','Science of Computer Programming',
         'Journal Of Systems And Software','Expert',
         'Automated Software Engineering','International Journal Of Software Engineering And Knowledge Engineering',
         'Foundations Of Software Engineering','Future Generation','Enterprise Information Systems',
         'Software Engineering Notes','Information and software technology',
         'International Conference on Service-Oriented Computing','International Requirements Engineering Conference',
         'Science of Computer Programming','Symposium on Applied Computing','Electronic Notes in Theoretical Computer Science',
         'International Workshop on Formal Aspects of Component Software','Formal Aspects of Computing',
         'Enterprise Distributed Object Computing Conference','IFAC Proceedings','International Conference on Formal Engineering Methods',
         'International CSI Computer Conference','European Symposium on Research in Computer Security',
         'International Conference on Fundamentals of Software Engineering','Applied Artificial Intelligence',
         'International Workshop on Engineering and Cybersecurity','Explainable Software for Cyber-Physical Systems',
         'European Conference on Software Architecture','Fundamental Approaches to Software Engineering',
         'international Workshop on Context-aware','International Conference on Distributed Computing Systems','IEEE Conference on Autonomic Computing',
         'IEEE International Conference on Robot and Human Interactive','International Conference on Formal Methods',
         'International Workshops on Foundations and Application','Journal of systems architecture','IEEE Software',
         'International Journal of Software Engineering','IEEE Access','Service-Oriented Computing','Requirements Engineering',
         'Information and software technology','arXiv','ACM Transactions on Cyber-Physical','Autonomic Computing',
         'Software Engineering for Self-Adaptive','Conference on Software Product',
         'Conference on Software Engineering','transactions on software engineering','CEUR','Workshop on Models','Models',
         'Automated software','Foundations of software','requirements','Cluster','Autonomic','Software engineering','Trusted Computing',
         'Intelligent Transport','Formal Method','Verification and Validation']

lw_searchfor=[]
lw_sources=[]

for s in searchfor:
    lw_searchfor.append(s.lower())
    
for s in sources:
    lw_sources.append(s.lower())
    
s_source = final_df_all_fine_grained[final_df_all_fine_grained['Source title'].str.contains('|'.join(lw_sources), na=False)]
    
# final_df_all_fine_grained['Source title'] = final_df_all_fine_grained['Source title'].str.lower() 
# final_df_all_fine_grained['Abstract'] = final_df_all_fine_grained['Abstract'].str.lower() 
   
s_1 = s_source[s_source.Abstract.str.contains('|'.join(lw_searchfor))]

writer = pd.ExcelWriter(s_path, engine='xlsxwriter',options={'strings_to_urls': False})
s_source.to_excel(writer)
writer.close()


