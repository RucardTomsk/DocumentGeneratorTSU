from py2neo import Graph
import config

class API(object):
	def __init__(self,name_database):
		self.graph = Graph(config.URL,auth=config.Account,name=name_database)

	def set_plan_text(self, guid,text):
		self.graph.run(f"MATCH (n:Plan) WHERE n.guid ='{guid}' SET n.ContentDiscipline = '{text}'")

	def set_plan_textHTML(self,guid,text):
		self.graph.run(f"MATCH (n:PlanHTML) WHERE n.guid ='{guid + '_html'}' SET n.ContentDiscipline = '{text}'")

	def get_plan_code_name(self):
		mas1 = self.graph.run(f"MATCH (n:Plan) RETURN COLLECT(n.guid)").evaluate()
		mas2 =self.graph.run(f"MATCH (n:Plan) RETURN COLLECT(n.Code)").evaluate()
		mas3 = self.graph.run(f"MATCH (n:Plan) RETURN COLLECT(n.Name)").evaluate()
		return mas1,mas2,mas3

	def get_guid_all_nodes(self):
		return self.graph.run(f"MATCH (n) RETURN COLLECT(n.guid)").evaluate()

	def get_node(self,guid):
		return self.graph.run(f"MATCH (n) WHERE n.guid = '{guid}' RETURN n").evaluate()

	def get_all_nodes_the_label(self,label):
		return self.graph.run(f"MATCH (n:{label}) RETURN COLLECT(n.guid)").evaluate()

	def get_node_name(self,guid):
		return self.graph.run(f"MATCH (n) WHERE n.guid = '{guid}' RETURN n.name").evaluate()

	def get_node_parents(self,guid):
		mas = self.graph.run(f"MATCH (p)-[]->(n) WHERE n.guid = '{guid}' RETURN COLLECT(p.guid)").evaluate()
		mas[0], mas[1] = mas[1], mas[0]
		return mas

	def get_node_parents_labels(self,guid,label):
		return self.graph.run(f"MATCH (p:{label})-[]->(n) WHERE n.guid = '{guid}' RETURN COLLECT(p.guid)").evaluate()

	def has_label_node(self,guid,label):
		if (self.graph.run(f"MATCH (n) WHERE n.guid = '{guid}' RETURN labels(n)").evaluate()[0] == label):
			return True
		else:
			return False

	def get_type_influence_node(self,guid_p,guid_n):
		if self.graph.run(f"MATCH (p)-[s]->(n) WHERE p.guid = '{guid_p}' AND n.guid = '{guid_n}' RETURN s.typeOfEvidence").evaluate() == "True":
			return True
		else:
			return False

	def get_degree_influence_node(self,guid_p,guid_n):
		return int(self.graph.run(f"MATCH (p)-[s]->(n) WHERE p.guid = '{guid_p}' AND n.guid = '{guid_n}' RETURN s.degreeOfEvidenceEnumValue").evaluate())

	def get_normalValue_node(self,guid):
		return int(self.graph.run(f"MATCH (n) WHERE n.guid = '{guid}' RETURN n.normalValue").evaluate())