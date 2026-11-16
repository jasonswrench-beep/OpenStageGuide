// A small design for the design desk's tests: two modules, three lessons, four objectives, activities from two models, assessments, and an evaluation plan.
import {cleanDesign} from '../../src/design-model.mjs';

export const raw=()=>({
  version:1,
  needs:{problem:'New supervisors handle conflict badly.',audience:'First-year supervisors',current:'They avoid it',desired:'They hold a short, calm conversation',causes:['practice','motivation'],constraints:'Two half days',success:'Fewer escalations',delivery:'In person'},
  objectives:[
    {id:'o1',text:'Learners will be able to explain the three steps of a coaching conversation.',level:''},
    {id:'o2',text:'Learners will be able to demonstrate a calm opening to a hard conversation.',level:'apply'},
    {id:'o3',text:'Learners will understand the escalation policy.',level:''},
    {id:'o4',text:'Learners will be able to justify when to escalate.',level:'evaluate'},
  ],
  models:['gagne','merrill'],
  activities:[
    {id:'a1',title:'Story of a hard conversation',model:'gagne',event:'Gain attention',minutes:10,objectives:['o1'],note:'Use a real case.'},
    {id:'a2',title:'Role play: the opening',model:'merrill',event:'Application: let them do it',minutes:25,objectives:['o2'],note:''},
    {id:'a3',title:'Gallery walk',model:'',event:'',minutes:15,objectives:[],note:''},
  ],
  assessments:[
    {id:'s1',title:'Skills check',kind:'demo',level:'remember',slots:0,objectives:['o2'],note:''},
    {id:'s2',title:'Short quiz',kind:'quiz',level:'understand',slots:2,objectives:['o1','o4'],note:'Five minutes.'},
    {id:'s3',title:'Unlinked check',kind:'self',level:'',slots:1,objectives:[],note:''},
  ],
  storyboard:[
    {id:'m1',title:'Talk first',lessons:[
      {id:'l1',title:'The coaching conversation',minutes:30,objectives:['o1'],activities:['a1'],assessments:['s2'],note:''},
      {id:'l2',title:'The calm opening',minutes:45,objectives:['o2'],activities:['a2'],assessments:['s1'],note:''},
    ]},
    {id:'m2',title:'Know the policy',lessons:[
      {id:'l3',title:'When to escalate',minutes:20,objectives:['o3','o4'],activities:[],assessments:[],note:''},
    ]},
  ],
  evaluation:{notes:'Ask again after three months.',measures:[
    {id:'e1',level:1,name:'End of day',kind:'rating',when:'At the end of the second half day',who:'Every learner',scale:'1, 2, 3, 4, 5',statements:['The session met its goals.','I can use what I learned.'],objectives:[],note:''},
    {id:'e2',level:2,name:'Escalation quiz',kind:'pre-post',when:'First and last hour',who:'Every learner',scale:'',statements:[],objectives:['o4'],note:''},
    {id:'e3',level:3,name:'Conversations observed',kind:'rating',when:'Six weeks later',who:'The learner\'s manager',scale:'Rarely, Sometimes, Often',statements:['Opens calmly.'],objectives:['o2'],note:''},
    {id:'e4',level:4,name:'Escalations per month',kind:'metric',when:'Before, and at six months',who:'The department',scale:'',statements:[],objectives:[],note:''},
  ]},
});
export const design=()=>cleanDesign(raw());
