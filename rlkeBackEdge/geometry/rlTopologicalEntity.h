#pragma once



class rlTopologicalEntity
{
	friend class rlTopologicalEntSet;
public:
	rlBaseClassTypeDef(emTopologicalEntity)

	rlId id() const;

protected:
	rlTopologicalEntity();
	virtual ~rlTopologicalEntity();

private:
	void _set(rlId id);

private:
	rlId _id;
};

