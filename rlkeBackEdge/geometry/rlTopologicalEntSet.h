#pragma once
#include "rl_id_manager.hpp"

class rlTopologicalEntity;
class rlTopologicalEntSet
{
public:
	template<class T>
	const T get(rlId key)const
	{
		std::map<rlId, rlTopologicalEntity*>::const_iterator it = _ptrs.find(key);
		if (it == _ptrs.end())
			return NULL;

		return dynamic_cast<T>(it->second);
	}


protected:
	rlTopologicalEntSet();
	virtual ~rlTopologicalEntSet();

	template<class T>
	T *__new()
	{
		rlTopologicalEntity *t = new T();
		t->_set(_id_mgr->id());
		_ptrs.insert(std::make_pair(t->id(), t));

		return dynamic_cast<T *>(t);
	}
	void __delete(rlTopologicalEntity *entity);

	void _insert(rlId id, rlTopologicalEntity *entity);
	rl_id_manager *id_mgr() const;
	const std::map<rlId, rlTopologicalEntity*> &ptrs() const;

private:
	void _clear();

private:
	std::map<rlId, rlTopologicalEntity*> _ptrs;
	rl_id_manager *_id_mgr;
};

