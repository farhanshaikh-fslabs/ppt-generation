from core.config import dynamodb
from boto3.dynamodb.conditions import Key, Attr
from core.config import DYNAMODB_SIMULATIONS_TABLE

def insert_data(table_name: str, data: dict) -> None:
    """Insert data into DynamoDB."""
    try:
        table = dynamodb.Table(table_name)

        # Simply put the item - DynamoDB will overwrite if it exists
        table.put_item(Item=data)
        print(f"Successfully inserted/updated data in DynamoDB table {table_name}")

    except Exception as e:
        print(f"Error inserting data into DynamoDB: {e}")
        raise

def get_company(table_name: str, company_name: str, entity_key: str):

    try:
        table = dynamodb.Table(table_name)
        response = table.get_item(Key={"company_name": company_name, "entity_key": entity_key})
        return response.get('Item', None)

    except Exception as e:
        print(f"Error retrieving from DynamoDB: {e}")
        return None

def get_all_companies(
    table_name: str,
    user_id: str | None = None,
    entity_key: str | None = None,
    *,
    index_name: str = "userid-index",
) -> list[dict]:
    try:
        table = dynamodb.Table(table_name)

        items: list[dict] = []
        last_evaluated_key = None

        # Prefer indexed query when user_id is available (partition key on GSI).
        if user_id:
            # Some deployments define `userid-index` with ONLY a partition key (`user_id`)
            # and no sort key. Adding a 2nd KeyCondition then fails validation.
            # So we always query by `user_id` and (optionally) filter by entity_key.
            key_condition = Key("user_id").eq(user_id)

            while True:
                kwargs: dict = {"IndexName": index_name, "KeyConditionExpression": key_condition}
                if entity_key:
                    kwargs["FilterExpression"] = Attr("entity_key").eq(entity_key)
                if last_evaluated_key:
                    kwargs["ExclusiveStartKey"] = last_evaluated_key
                response = table.query(**kwargs)
                items.extend(response.get("Items", []))
                last_evaluated_key = response.get("LastEvaluatedKey")
                if not last_evaluated_key:
                    break
            return items

        # If only the sort key filter comes in, we must scan + filter (can't query on sort key alone).
        scan_kwargs = {}
        if entity_key:
            scan_kwargs["FilterExpression"] = Attr("entity_key").eq(entity_key)

        while True:
            if last_evaluated_key:
                scan_kwargs["ExclusiveStartKey"] = last_evaluated_key
            response = table.scan(**scan_kwargs)
            items.extend(response.get("Items", []))
            last_evaluated_key = response.get("LastEvaluatedKey")
            if not last_evaluated_key:
                break
        return items
    except Exception as e:
        print(f"Error retrieving all companies from DynamoDB: {e}")
        return []

def get_all_companies_paginated(
    table_name: str,
    user_id: str | None = None,
    entity_key: str | None = None,
    *,
    limit: int = 50,
    exclusive_start_key: dict | None = None,
    index_name: str = "userid-index",
) -> tuple[list[dict], dict | None]:
    """
    Paginated version of get_all_companies.

    Returns: (items, last_evaluated_key)
    - If last_evaluated_key is not None, pass it back as exclusive_start_key to get next page.
    """
    try:
        table = dynamodb.Table(table_name)

        # Prefer indexed query when user_id is available (partition key on GSI).
        if user_id:
            kwargs: dict = {
                "IndexName": index_name,
                "KeyConditionExpression": Key("user_id").eq(user_id),
                "Limit": int(limit),
            }
            if entity_key:
                kwargs["FilterExpression"] = Attr("entity_key").eq(entity_key)
            if exclusive_start_key:
                kwargs["ExclusiveStartKey"] = exclusive_start_key
            response = table.query(**kwargs)
            return response.get("Items", []), response.get("LastEvaluatedKey")

        # If only the sort key filter comes in, we must scan + filter (can't query on sort key alone).
        scan_kwargs: dict = {"Limit": int(limit)}
        if entity_key:
            scan_kwargs["FilterExpression"] = Attr("entity_key").eq(entity_key)
        if exclusive_start_key:
            scan_kwargs["ExclusiveStartKey"] = exclusive_start_key
        response = table.scan(**scan_kwargs)
        return response.get("Items", []), response.get("LastEvaluatedKey")
    except Exception as e:
        print(f"Error retrieving paginated companies from DynamoDB: {e}")
        return [], None

def get_company_access(table_name: str, user_id: str, company_name: str) -> dict | None:
    """Check if a user_id + company_name entry exists in companies_access."""
    try:
        table = dynamodb.Table(table_name)
        response = table.get_item(Key={"user_id": user_id, "company_name": company_name})
        return response.get("Item", None)
    except Exception as e:
        print(f"Error retrieving from companies_access: {e}")
        return None


def get_companies_by_user(
    table_name: str,
    user_id: str,
    *,
    limit: int | None = None,
    exclusive_start_key: dict | None = None,
) -> tuple[list[dict], dict | None]:
    """Get company_name entries for a user from companies_access with optional pagination.

    Returns (items, last_evaluated_key).
    """
    try:
        table = dynamodb.Table(table_name)
        kwargs: dict = {"KeyConditionExpression": Key("user_id").eq(user_id)}
        if limit:
            kwargs["Limit"] = int(limit)
        if exclusive_start_key:
            kwargs["ExclusiveStartKey"] = exclusive_start_key
        response = table.query(**kwargs)
        return response.get("Items", []), response.get("LastEvaluatedKey")
    except Exception as e:
        print(f"Error querying companies_access by user: {e}")
        return [], None


def batch_get_companies(table_name: str, keys: list[dict]) -> list[dict]:
    """Fetch multiple companies in a single BatchGetItem call.

    Args:
        keys: list of {"company_name": ..., "entity_key": ...} dicts.
    DynamoDB limits BatchGetItem to 100 keys per request, so we chunk automatically.
    """
    if not keys:
        return []
    try:
        table = dynamodb.Table(table_name)
        items: list[dict] = []
        for i in range(0, len(keys), 100):
            chunk = keys[i : i + 100]
            response = dynamodb.batch_get_item(
                RequestItems={table_name: {"Keys": chunk}}
            )
            items.extend(response.get("Responses", {}).get(table_name, []))
            unprocessed = response.get("UnprocessedKeys", {}).get(table_name, {}).get("Keys", [])
            while unprocessed:
                response = dynamodb.batch_get_item(
                    RequestItems={table_name: {"Keys": unprocessed}}
                )
                items.extend(response.get("Responses", {}).get(table_name, []))
                unprocessed = response.get("UnprocessedKeys", {}).get(table_name, {}).get("Keys", [])
        return items
    except Exception as e:
        print(f"Error batch getting companies: {e}")
        return []


def update_data(key_value: dict, update_data: dict, table_name: str) -> None:
    """Update item in DynamoDB. Uses ExpressionAttributeNames so reserved words (e.g. status) are allowed."""
    try:
        table = dynamodb.Table(table_name)

        update_expressions = []
        expression_attribute_names = {}
        expression_attribute_values = {}

        for idx, (field, value) in enumerate(update_data.items()):
            name_placeholder = f"#n{idx}"
            value_placeholder = f":val{idx}"
            expression_attribute_names[name_placeholder] = field
            expression_attribute_values[value_placeholder] = value
            update_expressions.append(f"{name_placeholder} = {value_placeholder}")

        update_expression = "SET " + ", ".join(update_expressions)

        table.update_item(
            Key=key_value,
            UpdateExpression=update_expression,
            ExpressionAttributeNames=expression_attribute_names,
            ExpressionAttributeValues=expression_attribute_values,
        )

        print(f"Updated successfully.")

    except Exception as e:
        print(f"Error updating data in DynamoDB: {e}")
        raise

def get_simulation_data(table_name: str, simulation_id: str, record_type: str):
    try:
        table = dynamodb.Table(table_name)
        response = table.get_item(Key={"simulation_id": simulation_id, "record_type": record_type})
        return response.get('Item', None)
    except Exception as e:
        print(f"Error retrieving simulation from DynamoDB: {e}")
        return None


def get_simulation_by_id(table_name: str, simulation_id: str) -> dict:
    """Fetch all records for a simulation: metadata, state (strategy, summary, report), turns, judge feedbacks."""
    try:
        table = dynamodb.Table(table_name)
        response = table.query(KeyConditionExpression=Key("simulation_id").eq(simulation_id))
        items = response.get("Items", [])
        out = {"simulation_id": simulation_id, "metadata": None, "state": None, "turns": [], "judge_feedbacks": []}
        for item in items:
            rt = item.get("record_type", "")
            if rt == "metadata":
                out["metadata"] = item
            elif rt == "state":
                out["state"] = item
            elif rt.startswith("turn_") and rt != "turn_" and not rt.startswith("turn_judge"):
                try:
                    turn_num = int(rt.replace("turn_", ""))
                    out["turns"].append({"turn_id": turn_num, "record_type": rt, **item})
                except ValueError:
                    pass
            elif rt.startswith("judge_"):
                try:
                    after_turn = int(rt.replace("judge_", ""))
                    out["judge_feedbacks"].append({"after_turn_id": after_turn, "record_type": rt, **item})
                except ValueError:
                    pass
        out["turns"].sort(key=lambda x: x["turn_id"])
        out["judge_feedbacks"].sort(key=lambda x: x["after_turn_id"])
        return out
    except Exception as e:
        print(f"Error retrieving simulation by id from DynamoDB: {e}")
        return {"simulation_id": simulation_id, "metadata": None, "state": None, "turns": [], "judge_feedbacks": [], "error": str(e)}


def get_all_simulations(
    table_name: str,
    record_type: str = "metadata",
    user_id: str | None = None,
    *,
    index_name: str = "userid-index",
) -> list[dict]:
    try:
        table = dynamodb.Table(table_name)
        items: list[dict] = []
        last_evaluated_key = None

        if user_id:
            # Query user simulations via GSI, then filter to the desired record_type (metadata).
            while True:
                kwargs: dict = {
                    "IndexName": index_name,
                    "KeyConditionExpression": Key("user_id").eq(user_id),
                    "FilterExpression": Attr("record_type").eq(record_type),
                }
                if last_evaluated_key:
                    kwargs["ExclusiveStartKey"] = last_evaluated_key
                response = table.query(**kwargs)
                items.extend(response.get("Items", []))
                last_evaluated_key = response.get("LastEvaluatedKey")
                if not last_evaluated_key:
                    break
            return items

        # No user_id filter -> scan + filter by record_type.
        scan_kwargs: dict = {"FilterExpression": Attr("record_type").eq(record_type)}
        while True:
            if last_evaluated_key:
                scan_kwargs["ExclusiveStartKey"] = last_evaluated_key
            response = table.scan(**scan_kwargs)
            items.extend(response.get("Items", []))
            last_evaluated_key = response.get("LastEvaluatedKey")
            if not last_evaluated_key:
                break
        return items
    except Exception as e:
        print(f"Error retrieving all simulations from DynamoDB: {e}")
        return []

def get_all_simulations_paginated(
    table_name: str,
    record_type: str = "metadata",
    user_id: str | None = None,
    *,
    limit: int = 50,
    exclusive_start_key: dict | None = None,
) -> tuple[list[dict], dict | None]:
    try:
        table = dynamodb.Table(table_name)

        if user_id:
            filter_expr = Attr("record_type").eq(record_type) & Attr("user_id").eq(user_id)
        else:
            filter_expr = Attr("record_type").eq(record_type)

        collected = []
        last_evaluated_key = exclusive_start_key

        while len(collected) < limit:
            scan_kwargs: dict = {
                "FilterExpression": filter_expr,
                "Limit": limit,  # scan up to `limit` items per round
            }
            if last_evaluated_key:
                scan_kwargs["ExclusiveStartKey"] = last_evaluated_key

            response = table.scan(**scan_kwargs)
            collected.extend(response.get("Items", []))
            last_evaluated_key = response.get("LastEvaluatedKey")

           

            # No more pages left
            if not last_evaluated_key:
                break

        # Attach the corresponding "state" record for each simulation_id
        # (each metadata record is expected to contain simulation_id).
        for item in collected:
            sim_id = item.get("simulation_id")
            if not sim_id:
                item["state"] = None
                continue

            state_item = table.get_item(
                Key={"simulation_id": sim_id, "record_type": "state"}
            ).get("Item")
            item["state"] = state_item
        # If we collected more than needed, trim and retain cursor for next page
        if len(collected) > limit:
            # Can't resume mid-batch easily, so just return trimmed without cursor
            return collected[:limit], last_evaluated_key
        
        return collected, last_evaluated_key

    except Exception as e:
        print(f"Error retrieving paginated simulations from DynamoDB: {e}")
        return [], None

if __name__ == "__main__":
    print(get_all_simulations(DYNAMODB_SIMULATIONS_TABLE, "metadata"))