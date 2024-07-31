import { Storage, StoreItems } from 'botbuilder-core';
import { CosmosClient } from '@azure/cosmos';

export class CosmosStorage implements Storage {
	private client: CosmosClient;
	private databaseId: string;
	private containerId: string;

	constructor(endpoint: string, key: string, databaseId: string, containerId: string) {
		this.client = new CosmosClient({ endpoint, key });
		this.databaseId = databaseId;
		this.containerId = containerId;
	}

	async read(keys: string[]): Promise<StoreItems> {
		if (!keys || keys.length === 0) {
			return {};
		}

		const container = this.client.database(this.databaseId).container(this.containerId);
		const storeItems: StoreItems = {};

		for (const key of keys) {
            const sanitizedKey = this.sanitizeKey(key);
			const { resource: item } = await container.item(sanitizedKey, sanitizedKey).read();
			if (item) {
				storeItems[key] = item;
			}
		}

		return storeItems;
	}

	async write(changes: StoreItems): Promise<void> {
		const container = this.client.database(this.databaseId).container(this.containerId);
		const operations = [];

		for (const key in changes) {
			if (changes.hasOwnProperty(key)) {
				const item = changes[key];
                const sanitizedKey = this.sanitizeKey(key);
                const oldItem = await container.item(sanitizedKey, sanitizedKey).read();
                if (!oldItem.resource || item.eTag === '*' || !item.eTag) {
                    item.id = sanitizedKey; // Ensure the item has an id
                    operations.push(container.items.upsert(item));
                } else {
                    // parse old item to get etag

                    if (oldItem.resource.eTag === item.eTag) {
                        item.id = sanitizedKey; // Ensure the item has an id
                        operations.push(container.items.upsert(item));
                    } else {
                        throw new Error(`Storage: error writing "${key}" due to eTag conflict.`);
                    }
                }
			}
		}

		await Promise.all(operations);
	}

	async delete(keys: string[]): Promise<void> {
		if (!keys || keys.length === 0) {
			return;
		}

		const container = this.client.database(this.databaseId).container(this.containerId);
		const operations = [];

		for (const key of keys) {
			operations.push(container.item(this.sanitizeKey(key), this.sanitizeKey(key)).delete());
		}

		await Promise.all(operations);
	}

    private sanitizeKey(key: string): string {
        if (!key || !key.length) {
            throw new Error('Please provide a non-empty key');
        }
    
        const sanitized = key.replace(/\//g, '');
        return sanitized.substring(0, 255);
    }
}